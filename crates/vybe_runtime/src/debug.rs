use crate::chunk::Chunk;
use crate::opcode::{Op, OperandFormat, read_leb_u32, read_leb_u64};

/// Disassemble a chunk, resolving `GLOBAL_GET`/`GLOBAL_SET` operands against
/// the module's global table.
///
/// The table needs no parameter: `Chunk::globals` is the module-level index
/// space, shared by every chunk (`globals::normalize_global_table`). A chunk
/// that carries a globalidx can therefore always say what it means. The earlier
/// `disassemble_with(chunk, globals)` split existed only because the table sat
/// on chunk 0 alone, which forced every caller to fetch it and hand it back —
/// and made a caller edit the price of naming a global.
pub fn disassemble(chunk: &Chunk) -> String {
    let mut out = String::new();
    out.push_str(&format!("== {} ==\n", chunk.name));
    let mut offset = 0;
    while offset < chunk.code.len() {
        let (text, next) = disassemble_instruction(chunk, offset);
        out.push_str(&format!("{:04}  {}\n", offset, text));
        offset = next;
    }
    out
}

pub fn disassemble_instruction(chunk: &Chunk, offset: usize) -> (String, usize) {
    disassemble_instruction_inner(chunk, &chunk.globals, offset)
}

fn disassemble_instruction_inner(
    chunk: &Chunk,
    globals: &[String],
    offset: usize,
) -> (String, usize) {
    if offset + 3 >= chunk.code.len() {
        return ("TRUNCATED".into(), chunk.code.len());
    }

    let group = ((chunk.code[offset] as u16) << 8) | chunk.code[offset + 1] as u16;
    let sub = ((chunk.code[offset + 2] as u16) << 8) | chunk.code[offset + 3] as u16;
    let op = match Op::decode(group, sub) {
        Some(op) => op,
        None => {
            return (
                format!("UNKNOWN(0x{:04X} 0x{:04X})", group, sub),
                offset + 4,
            );
        }
    };

    let operand_start = offset + 4;
    let name = op.wasm_name();

    match op.operand_format() {
        OperandFormat::None => (format!("{}", name), operand_start),
        OperandFormat::U8 => {
            let arg = chunk.code.get(operand_start).copied().unwrap_or(0);
            (format!("{} {}", name, arg), operand_start + 1)
        }
        OperandFormat::U8_U8 => {
            let a = chunk.code.get(operand_start).copied().unwrap_or(0);
            let b = chunk.code.get(operand_start + 1).copied().unwrap_or(0);
            (format!("{} {} {}", name, a, b), operand_start + 2)
        }
        OperandFormat::U8_U8_U8 => {
            let a = chunk.code.get(operand_start).copied().unwrap_or(0);
            let b = chunk.code.get(operand_start + 1).copied().unwrap_or(0);
            let c = chunk.code.get(operand_start + 2).copied().unwrap_or(0);
            (format!("{} {} {} {}", name, a, b, c), operand_start + 3)
        }
        OperandFormat::U16 => {
            let idx = chunk.read_u16(operand_start);
            // Ops whose U16 operand indexes the constant pool for a key/name/
            // value string. Resolving it is essential for reading property
            // access (`struct.get "prototype"`), global access (`global.get
            // "B"`), and literals — without it a disassembly of e.g. a super
            // chain is unreadable (just `struct.get 1/2/3`). LOCAL_GET/SET use
            // a *slot* index, not the constant pool, so they are excluded.
            // ⚠ `GLOBAL_GET`/`GLOBAL_SET` carry a GLOBALIDX, not a constant
            // index — they index the module's global table. Resolving them
            // against the constant pool (as this did) produces a label only
            // when the two numbers coincide, which reads as a correct name at
            // low indices and nothing at high ones. That cost real debugging
            // time: it looked like a duplicate-name bug in the table when the
            // table was fine.
            let references_constant = matches!(
                op,
                Op::STRUCT_GET | Op::STRUCT_GET_S | Op::STRUCT_GET_U | Op::STRUCT_SET
            );
            let is_global = matches!(op, Op::GLOBAL_GET | Op::GLOBAL_SET);
            let extra = if is_global {
                match globals.get(idx as usize) {
                    Some(name) => format!(" ({})", name),
                    None if globals.is_empty() => String::new(),
                    None => " (OUT-OF-RANGE)".to_string(),
                }
            } else if references_constant && (idx as usize) < chunk.constants.len() {
                format!(" ({})", chunk.constants[idx as usize])
            } else {
                String::new()
            };
            (format!("{} {}{}", name, idx, extra), operand_start + 2)
        }
        OperandFormat::I16 => {
            let off = chunk.read_i16(operand_start);
            let target = (operand_start as i64 + 2 + off as i64) as usize;
            (format!("{} {} -> {}", name, off, target), operand_start + 2)
        }
        OperandFormat::U16_U8 => {
            // call_import: u16 fn_index, u8 arg_count
            let fn_idx = chunk.read_u16(operand_start);
            let argc = chunk.code.get(operand_start + 2).copied().unwrap_or(0);
            // Resolve the import name from this chunk's own table — mismatched
            // per-chunk import indices are a recurring bug class and invisible
            // without the name.
            let label = match chunk.imports.get(fn_idx as usize) {
                Some(imp) => format!(" ({}:{})", imp.module, imp.name),
                None => " (OUT-OF-RANGE)".to_string(),
            };
            (
                format!("CallHost fn={} argc={}{}", fn_idx, argc, label),
                operand_start + 3,
            )
        }
        // ⛔ THIS FORMAT IS SHARED, SO IT CANNOT SPEAK ONE INSTRUCTION'S
        // VOCABULARY. The arm below already carried this scar once — it
        // hardcoded the retired `try_start`'s "catch/finally" wording and
        // mislabelled every op that later adopted the format. The same thing
        // happened again here: `br_on_cast_desc_eq` was rendered as
        // `table=… tag=… argc=…`, which reads like a `call_indirect_with_tag`
        // and hides the label depth — in a disassembly being used to CHECK
        // that instruction's immediates.
        OperandFormat::U16_U16_U8
            if op == Op::BR_ON_CAST_DESC_EQ || op == Op::BR_ON_CAST_DESC_EQ_FAIL =>
        {
            // `br_on_cast_desc_eq $l ht_1 ht_2`: the two u16s are constant
            // indices naming ht_2 (target) then ht_1 (source), and the u8 is
            // the label depth.
            let to = chunk.read_u16(operand_start);
            let from = chunk.read_u16(operand_start + 2);
            let depth = chunk.code.get(operand_start + 4).copied().unwrap_or(0);
            let spell = |idx: u16| match chunk.constants.get(idx as usize) {
                Some(crate::Value::String(s)) => s.to_string(),
                _ => idx.to_string(),
            };
            (
                format!(
                    "{} depth={} from={} to={}",
                    name,
                    depth,
                    spell(from),
                    spell(to)
                ),
                operand_start + 5,
            )
        }
        // `call_indirect_with_tag $table $tag argc` — the tag immediate NAMES
        // the tag (a constant-pool index), as `call_with_tag`'s does.
        OperandFormat::U16_U16_U8 => {
            let table = chunk.read_u16(operand_start);
            let tag = chunk.read_u16(operand_start + 2);
            let argc = chunk.code.get(operand_start + 4).copied().unwrap_or(0);
            let tag_name = match chunk.constants.get(tag as usize) {
                Some(crate::Value::String(s)) => format!(" ({s})"),
                _ => String::new(),
            };
            (
                format!("{} table={} tag={}{} argc={}", name, table, tag, tag_name, argc),
                operand_start + 5,
            )
        }
        OperandFormat::U16_U16 => {
            // Every U16_U16 op is a WASM GC instruction carrying a type
            // immediate plus a second one. (This arm used to hardcode the
            // retired `try_start`'s "catch/finally" text, which mislabelled
            // all of them once the struct ops widened to two immediates —
            // `try_table` is 0x00 0x1F and does not use this format.)
            let a = chunk.read_u16(operand_start);
            let b = chunk.read_u16(operand_start + 2);
            let text = match op {
                // `(typeidx, fieldidx)`. typeidx 0 is the dynamic form, where
                // the second immediate is a constant-pool property NAME —
                // resolving it is what makes property access readable
                // (`struct.get 3 (whoami)`), as documented in
                // `documentation/debugging.md`. A non-zero typeidx makes it a
                // spec fieldidx into indexed storage, which has no name.
                Op::STRUCT_GET | Op::STRUCT_GET_S | Op::STRUCT_GET_U | Op::STRUCT_SET => {
                    if a == 0 {
                        let named = (b as usize) < chunk.constants.len();
                        if named {
                            format!("{} {} ({})", name, b, chunk.constants[b as usize])
                        } else {
                            format!("{} {}", name, b)
                        }
                    } else {
                        format!("{} type={} field={}", name, a, b)
                    }
                }
                // `(typeidx, count)` — typeidx 0 builds a dynamic object from
                // `count` key/value pairs; non-zero is the spec `struct.new $t`.
                Op::STRUCT_NEW => format!("{} type={} count={}", name, a, b),
                _ => format!("{} type={} {}", name, a, b),
            };
            (text, operand_start + 4)
        }
        OperandFormat::U16_I16 => {
            // br_on_cast: u16 type_name + i16 offset
            let type_idx = chunk.read_u16(operand_start);
            let off = chunk.read_i16(operand_start + 2);
            (
                format!("{} type={} offset={}", name, type_idx, off),
                operand_start + 4,
            )
        }
        OperandFormat::U32Leb => {
            let mut next = operand_start;
            let value = read_leb_u32(&chunk.code, &mut next);
            (format!("{} {}", name, value), next)
        }
        OperandFormat::U32Leb_U32Leb => {
            let mut next = operand_start;
            let a = read_leb_u32(&chunk.code, &mut next);
            let b = read_leb_u32(&chunk.code, &mut next);
            (format!("{} {} {}", name, a, b), next)
        }
        OperandFormat::MemArg => {
            let mut next = operand_start;
            let align = read_leb_u32(&chunk.code, &mut next);
            let offset = read_leb_u32(&chunk.code, &mut next);
            let memidx = if align & 0x40 != 0 {
                Some(read_leb_u32(&chunk.code, &mut next))
            } else {
                None
            };
            let mem = memidx.map(|idx| format!(" mem={idx}")).unwrap_or_default();
            (
                format!("{} align={} offset={}{}", name, align, offset, mem),
                next,
            )
        }
        OperandFormat::MemArg64 => {
            let mut next = operand_start;
            let align = read_leb_u32(&chunk.code, &mut next);
            let offset = read_leb_u64(&chunk.code, &mut next);
            let memidx = if align & 0x40 != 0 {
                Some(read_leb_u32(&chunk.code, &mut next))
            } else {
                None
            };
            let mem = memidx.map(|idx| format!(" mem={idx}")).unwrap_or_default();
            (
                format!("{} align={} offset={}{}", name, align, offset, mem),
                next,
            )
        }
        // SIMD memory op: OPTIONAL marker-tagged memarg (0x80 bit on the
        // first LEB signals presence; absent = zero operand bytes).
        OperandFormat::SimdMemArg => {
            let size = vybe_runtime_simd_memarg_render(&chunk.code, operand_start);
            (format!("{}{}", name, size.0), size.1)
        }
        // SIMD lane mem op: the same optional memarg, then the lane byte.
        OperandFormat::MemLane => {
            let (memarg_txt, after_memarg) =
                vybe_runtime_simd_memarg_render(&chunk.code, operand_start);
            let lane = chunk.code.get(after_memarg).copied().unwrap_or(0);
            (
                format!("{}{} lane={}", name, memarg_txt, lane),
                after_memarg + 1,
            )
        }
        OperandFormat::Closure => {
            // ref_func: u16 func_index, u8 upvalue_count, then
            // (u8 is_local + u16 index) descriptors
            let func_idx = chunk.read_u16(operand_start);
            let uv_count =
                (chunk.code.get(operand_start + 2).copied().unwrap_or(0) & 0x7f) as usize;
            let total = 3 + uv_count * 3;
            (
                format!("Closure func={} upvalues={}", func_idx, uv_count),
                operand_start + total,
            )
        }
        OperandFormat::BrTable => {
            let mut next = operand_start;
            let count = read_leb_u32(&chunk.code, &mut next) as usize;
            let mut labels = Vec::with_capacity(count);
            for _ in 0..count {
                labels.push(read_leb_u32(&chunk.code, &mut next));
            }
            let default = read_leb_u32(&chunk.code, &mut next);
            (
                format!("br_table labels={:?} default={}", labels, default),
                next,
            )
        }
        OperandFormat::TryTable => {
            // try_table: u8 params, u8 results (the spec blocktype), u16 count,
            // then count × (u8 kind + u16 tag + u16 label)
            let count = ((chunk.code.get(operand_start + 2).copied().unwrap_or(0) as usize) << 8)
                | chunk.code.get(operand_start + 3).copied().unwrap_or(0) as usize;
            let total = 4 + count * 5;
            let mut clauses = Vec::with_capacity(count);
            for i in 0..count {
                let base = operand_start + 4 + i * 5;
                let kind = chunk.code.get(base).copied().unwrap_or(0);
                let tag = chunk.read_u16(base + 1);
                let label = chunk.read_u16(base + 3);
                let name = match kind {
                    0 => format!("catch tag={tag} → label {label}"),
                    1 => format!("catch_ref tag={tag} → label {label}"),
                    2 => format!("catch_all → label {label}"),
                    3 => format!("catch_all_ref → label {label}"),
                    k => format!("kind{k}? → label {label}"),
                };
                clauses.push(name);
            }
            let params = chunk.code.get(operand_start).copied().unwrap_or(0);
            let results = chunk.code.get(operand_start + 1).copied().unwrap_or(0);
            (
                format!(
                    "try_table (p{params} r{results}) [{}]",
                    clauses.join(", ")
                ),
                operand_start + total,
            )
        }
        OperandFormat::V128Const => (format!("v128.const [16 bytes]"), operand_start + 16),
        OperandFormat::Shuffle => (format!("i8x16.shuffle [16 indices]"), operand_start + 16),
        OperandFormat::SlI32 => {
            let mut next = operand_start;
            let val = crate::opcode::read_leb_i32(&chunk.code, &mut next);
            (format!("{} {}", name, val), next)
        }
        OperandFormat::SlI64 => {
            let mut next = operand_start;
            let val = crate::opcode::read_leb_i64(&chunk.code, &mut next);
            (format!("{} {}", name, val), next)
        }
        OperandFormat::RawF32 => {
            let bytes: [u8; 4] = chunk.code[operand_start..operand_start + 4]
                .try_into()
                .unwrap_or([0; 4]);
            let val = f32::from_le_bytes(bytes);
            (format!("{} {}", name, val), operand_start + 4)
        }
        OperandFormat::RawF64 => {
            let bytes: [u8; 8] = chunk.code[operand_start..operand_start + 8]
                .try_into()
                .unwrap_or([0; 8]);
            let val = f64::from_le_bytes(bytes);
            (format!("{} {}", name, val), operand_start + 8)
        }
    }
}

/// Render the optional marker-tagged SIMD memarg (see `SimdMemArg`). Returns
/// the rendered suffix (empty when absent) and the offset just past it.
fn vybe_runtime_simd_memarg_render(code: &[u8], operand_start: usize) -> (String, usize) {
    let mut ip = operand_start;
    let align = read_leb_u32(code, &mut ip);
    if align & 0x80 == 0 {
        return (String::new(), operand_start);
    }
    let offset = read_leb_u64(code, &mut ip);
    let memidx = if align & 0x40 != 0 {
        Some(read_leb_u32(code, &mut ip))
    } else {
        None
    };
    let mem = memidx.map(|idx| format!(" mem={idx}")).unwrap_or_default();
    (
        format!(" align={} offset={}{}", align & !0x1C0, offset, mem),
        ip,
    )
}

// ── Bytecode verifier ────────────────────────────────────────────────────

/// One structural defect found by [`verify_chunk`].
#[derive(Debug, Clone)]
pub struct VerifyIssue {
    /// Byte offset the problem was found at.
    pub offset: usize,
    pub what: String,
}

/// Check a chunk's two structural invariants and report every violation.
///
/// 1. **Every instruction sits on the opcode grid.** Opcodes are always 4
///    bytes plus a format-determined operand, so walking from 0 must decode
///    cleanly to the end. A byte that fails to decode means something earlier
///    mis-declared its operand width.
/// 2. **Every jump lands on an instruction start.** A target that falls inside
///    another instruction is the classic symptom — the VM decodes the tail of
///    one opcode as the head of another and reports a nonsense opcode like
///    `0x0B00 0x0000` (the last byte of an `end`, plus the next three).
///
/// These are the bugs that cost the most to find by hand: the failure surfaces
/// far from the emitter that caused it, and only for input shapes that happen
/// to change a body's length.
pub fn verify_chunk(chunk: &Chunk) -> Vec<VerifyIssue> {
    let code = &chunk.code;
    let mut issues = Vec::new();
    let mut starts = std::collections::HashSet::new();
    let mut jumps: Vec<(usize, i64, &'static str)> = Vec::new();

    let mut ip = 0usize;
    while ip + 3 < code.len() {
        starts.insert(ip);
        let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
        let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
        let Some(op) = Op::decode(group, sub) else {
            issues.push(VerifyIssue {
                offset: ip,
                what: format!(
                    "does not decode as an opcode (0x{:04X} 0x{:04X}) — the previous \
                     instruction's operand width is wrong",
                    group, sub
                ),
            });
            break;
        };
        let operand_start = ip + 4;
        let size = op.operand_format().size_in(code, operand_start);

        // Relative branch offsets: `i16` immediately after the opcode.
        if matches!(op.operand_format(), crate::opcode::OperandFormat::I16)
            && operand_start + 1 < code.len()
        {
            let rel = i16::from_be_bytes([code[operand_start], code[operand_start + 1]]) as i64;
            jumps.push((
                operand_start,
                operand_start as i64 + 2 + rel,
                op.wasm_name(),
            ));
        }
        // `try_table` catch targets: u8 count, then per clause
        // (u8 kind, u16 tag, i16 catch offset relative to just past itself).
        if matches!(op.operand_format(), crate::opcode::OperandFormat::TryTable) {
            let count = code.get(operand_start).copied().unwrap_or(0) as usize;
            for c in 0..count {
                let pos = operand_start + 1 + c * 5 + 3;
                if pos + 1 < code.len() {
                    let rel = i16::from_be_bytes([code[pos], code[pos + 1]]) as i64;
                    jumps.push((pos, pos as i64 + 2 + rel, "try_table catch"));
                }
            }
        }
        ip = operand_start + size;
    }

    for (at, target, what) in jumps {
        if target < 0 || target as usize > code.len() {
            issues.push(VerifyIssue {
                offset: at,
                what: format!("{what} target {target} is outside the chunk"),
            });
        } else if target as usize != code.len() && !starts.contains(&(target as usize)) {
            issues.push(VerifyIssue {
                offset: at,
                what: format!("{what} target {target} is not an instruction start"),
            });
        }
    }
    issues
}
