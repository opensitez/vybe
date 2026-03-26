use crate::chunk::Chunk;
use crate::opcode::Op;

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
    let byte = chunk.code[offset];
    let op = match Op::from_byte(byte) {
        Some(op) => op,
        None => return (format!("UNKNOWN({})", byte), offset + 1),
    };

    match op {
        // No operands
        Op::Pop | Op::Dup |
        Op::AddF | Op::SubF | Op::MulF | Op::DivF | Op::ModF | Op::NegF |
        Op::AddI | Op::SubI | Op::MulI |
        Op::Concat |
        Op::BitAnd | Op::BitOr | Op::BitXor | Op::BitNot | Op::Shl | Op::Shr | Op::UShr |
        Op::CmpEq | Op::CmpNe | Op::CmpLtF | Op::CmpGtF | Op::CmpLeF | Op::CmpGeF |
        Op::CmpLtS | Op::CmpGtS |
        Op::BoolNot |
        Op::Return | Op::Halt |
        Op::PushNull | Op::PushTrue | Op::PushFalse |
        Op::PushI32Zero | Op::PushI32One | Op::PushF64Zero |
        Op::GetIndex | Op::SetIndex |
        Op::IsNull | Op::IsString | Op::IsNumber | Op::IsBool | Op::IsObject | Op::IsFunction |
        Op::ToF64 | Op::ToI32 |
        Op::TryEnd | Op::Throw |
        Op::Inherit | Op::GetIterator | Op::IterNext | Op::Spread => {
            (format!("{:?}", op), offset + 1)
        }

        // u16 operand
        Op::Const | Op::GetLocal | Op::SetLocal | Op::GetGlobal | Op::SetGlobal |
        Op::GetProp | Op::SetProp | Op::NewObject | Op::NewArray | Op::Class | Op::Method => {
            let idx = chunk.read_u16(offset + 1);
            let extra = if op == Op::Const && (idx as usize) < chunk.constants.len() {
                format!(" ({})", chunk.constants[idx as usize])
            } else {
                String::new()
            };
            (format!("{:?} {}{}", op, idx, extra), offset + 3)
        }

        // u8 operand
        Op::GetUpvalue | Op::SetUpvalue | Op::Call | Op::StrConcat => {
            let arg = chunk.code[offset + 1];
            (format!("{:?} {}", op, arg), offset + 2)
        }

        // i16 offset
        Op::Jump | Op::JumpIfFalse | Op::JumpIfTrue | Op::JumpIfNull => {
            let off = chunk.read_i16(offset + 1);
            let target = (offset as i64 + 3 + off as i64) as usize;
            (format!("{:?} {} -> {}", op, off, target), offset + 3)
        }

        // CallHost: u16 fn_index, u8 arg_count
        Op::CallHost => {
            let fn_idx = chunk.read_u16(offset + 1);
            let argc = chunk.code[offset + 3];
            (format!("CallHost fn={} argc={}", fn_idx, argc), offset + 4)
        }

        // Closure: u16 func_index, u8 upvalue_count, then pairs
        Op::Closure => {
            let func_idx = chunk.read_u16(offset + 1);
            let uv_count = chunk.code[offset + 3] as usize;
            let total = 4 + uv_count * 2;
            (format!("Closure func={} upvalues={}", func_idx, uv_count), offset + total)
        }

        // TryStart: u16 catch, u16 finally
        Op::TryStart => {
            let catch = chunk.read_u16(offset + 1);
            let finally = chunk.read_u16(offset + 3);
            (format!("TryStart catch={} finally={}", catch, finally), offset + 5)
        }
    }
}
