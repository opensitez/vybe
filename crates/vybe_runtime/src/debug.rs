use crate::chunk::Chunk;
use crate::opcode::{Op, OperandFormat, read_leb_u32, read_leb_u64};

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
            let references_constant = matches!(
                op,
                Op::GLOBAL_GET
                    | Op::GLOBAL_SET
                    | Op::STRUCT_GET
                    | Op::STRUCT_GET_S
                    | Op::STRUCT_GET_U
                    | Op::STRUCT_SET
            );
            let extra = if references_constant && (idx as usize) < chunk.constants.len() {
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
                None => " (OUT-OF-RANGE)".to_string() };
            (
                format!("CallHost fn={} argc={}{}", fn_idx, argc, label),
                operand_start + 3,
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
                _ => format!("{} type={} {}", name, a, b) };
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
        // SIMD lane mem op: a single lane-index byte (the VM's optional-memarg
        // peek never consumes a byte since lane indices are < 0x80).
        OperandFormat::MemLane => {
            let lane = chunk.code.get(operand_start).copied().unwrap_or(0);
            (format!("{} lane={}", name, lane), operand_start + 1)
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
            // try_table: u8 count, then count × (u8 kind + u16 tag + u16 offset)
            let count = chunk.code.get(operand_start).copied().unwrap_or(0) as usize;
            let total = 1 + count * 5;
            let mut clauses = Vec::with_capacity(count);
            for i in 0..count {
                let base = operand_start + 1 + i * 5;
                let kind = chunk.code.get(base).copied().unwrap_or(0);
                let tag = chunk.read_u16(base + 1);
                let name = match kind {
                    0 => format!("catch tag={tag}"),
                    1 => format!("catch_ref tag={tag}"),
                    2 => "catch_all".to_string(),
                    3 => "catch_all_ref".to_string(),
                    k => format!("kind{k}?") };
                clauses.push(name);
            }
            (
                format!("try_table [{}]", clauses.join(", ")),
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

// ── Bytecode verifier ────────────────────────────────────────────────────

/// One structural defect found by [`verify_chunk`].
#[derive(Debug, Clone)]
pub struct VerifyIssue {
    /// Byte offset the problem was found at.
    pub offset: usize,
    pub what: String }

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
                ) });
            break;
        };
        let operand_start = ip + 4;
        let size = op.operand_format().size_in(code, operand_start);

        // Relative branch offsets: `i16` immediately after the opcode.
        if matches!(op.operand_format(), crate::opcode::OperandFormat::I16) && operand_start + 1 < code.len() {
            let rel = i16::from_be_bytes([code[operand_start], code[operand_start + 1]]) as i64;
            jumps.push((operand_start, operand_start as i64 + 2 + rel, op.wasm_name()));
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
                what: format!("{what} target {target} is outside the chunk") });
        } else if target as usize != code.len() && !starts.contains(&(target as usize)) {
            issues.push(VerifyIssue {
                offset: at,
                what: format!("{what} target {target} is not an instruction start") });
        }
    }
    issues
}
