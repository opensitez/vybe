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

fn disassemble_instruction(chunk: &Chunk, offset: usize) -> (String, usize) {
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
        OperandFormat::U16 => {
            let idx = chunk.read_u16(operand_start);
            let extra = if op == Op::CONST && (idx as usize) < chunk.constants.len() {
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
        OperandFormat::U16_U16 => {
            // try_start: u16 catch, u16 finally
            let a = chunk.read_u16(operand_start);
            let b = chunk.read_u16(operand_start + 2);
            (
                format!("TryStart catch={} finally={}", a, b),
                operand_start + 4,
            )
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
        OperandFormat::Closure => {
            // ref_func: u16 func_index, u8 upvalue_count, then pairs
            let func_idx = chunk.read_u16(operand_start);
            let uv_count = chunk.code.get(operand_start + 2).copied().unwrap_or(0) as usize;
            let total = 3 + uv_count * 2;
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
            // try_table: u8 count, then count × (u8 tag + u16 offset)
            let count = chunk.code.get(operand_start).copied().unwrap_or(0) as usize;
            let total = 1 + count * 3;
            (
                format!("try_table handlers={}", count),
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
