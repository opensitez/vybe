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
        Op::drop | Op::dup |
        Op::f64_add | Op::f64_sub | Op::f64_mul | Op::f64_div | Op::f64_mod | Op::f64_neg |
        Op::i32_add | Op::i32_sub | Op::i32_mul |
        Op::str_concat |
        Op::i32_and | Op::i32_or | Op::i32_xor | Op::i32_not | Op::i32_shl | Op::i32_shr_s | Op::i32_shr_u |
        Op::eq | Op::ne | Op::f64_lt | Op::f64_gt | Op::f64_le | Op::f64_ge |
        Op::str_lt | Op::str_gt |
        Op::bool_not |
        Op::dyn_add | Op::dyn_eq | Op::dyn_ne | Op::dyn_lt | Op::dyn_gt |
        Op::dyn_le | Op::dyn_ge | Op::dyn_neg | Op::dyn_not | Op::dyn_to_bool |
        Op::r#await | Op::set_timer |
        Op::r#return | Op::halt |
        Op::null | Op::r#true | Op::r#false |
        Op::i32_const_0 | Op::i32_const_1 | Op::f64_const_0 |
        Op::array_get | Op::array_set |
        Op::ref_is_null | Op::ref_is_string | Op::ref_is_number | Op::ref_is_bool | Op::ref_is_object | Op::ref_is_func |
        Op::f64_from_i32 | Op::i32_from_f64 |
        Op::try_end | Op::throw |
        Op::inherit | Op::iter_get | Op::iter_next | Op::spread |
        Op::memory_size | Op::end | Op::unpack => {
            (format!("{:?}", op), offset + 1)
        }

        // u16 operand
        Op::r#const | Op::local_get | Op::local_set | Op::global_get | Op::global_set |
        Op::struct_get | Op::struct_set | Op::struct_new | Op::array_new | Op::class_new | Op::method_def |
        Op::block | Op::r#loop | Op::memory_grow | Op::canon_lift | Op::canon_lower | Op::ref_test => {
            let idx = chunk.read_u16(offset + 1);
            let extra = if op == Op::r#const && (idx as usize) < chunk.constants.len() {
                format!(" ({})", chunk.constants[idx as usize])
            } else {
                String::new()
            };
            (format!("{:?} {}{}", op, idx, extra), offset + 3)
        }

        // u8 operand
        Op::upvalue_get | Op::upvalue_set | Op::call | Op::str_concat_n |
        Op::return_call | Op::call_indirect | Op::pack |
        Op::br_label | Op::br_if_label => {
            let arg = chunk.code[offset + 1];
            (format!("{:?} {}", op, arg), offset + 2)
        }

        // i16 offset
        Op::br | Op::br_if_false | Op::br_if_true | Op::br_if_null => {
            let off = chunk.read_i16(offset + 1);
            let target = (offset as i64 + 3 + off as i64) as usize;
            (format!("{:?} {} -> {}", op, off, target), offset + 3)
        }

        // CallHost: u16 fn_index, u8 arg_count
        Op::call_import => {
            let fn_idx = chunk.read_u16(offset + 1);
            let argc = chunk.code[offset + 3];
            (format!("CallHost fn={} argc={}", fn_idx, argc), offset + 4)
        }

        // Closure: u16 func_index, u8 upvalue_count, then pairs
        Op::ref_func => {
            let func_idx = chunk.read_u16(offset + 1);
            let uv_count = chunk.code[offset + 3] as usize;
            let total = 4 + uv_count * 2;
            (format!("Closure func={} upvalues={}", func_idx, uv_count), offset + total)
        }

        // Memory load/store — no operand, just opcode (addr on stack)
        Op::i32_load | Op::i32_store | Op::i64_load | Op::i64_store |
        Op::f64_load | Op::f64_store | Op::i32_load8_u | Op::i32_store8 => {
            (format!("{:?}", op), offset + 1)
        }

        // br_table: u8 count, u8 default, then count × u8 labels
        Op::br_table => {
            let count = chunk.code[offset + 1] as usize;
            let default = chunk.code[offset + 2];
            let total = 3 + count;
            (format!("br_table count={} default={}", count, default), offset + total)
        }

        // TryStart: u16 catch, u16 finally
        Op::try_start => {
            let catch = chunk.read_u16(offset + 1);
            let finally = chunk.read_u16(offset + 3);
            (format!("TryStart catch={} finally={}", catch, finally), offset + 5)
        }
    }
}
