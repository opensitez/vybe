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
    let (op, base) = if byte == 0xFE && offset + 1 < chunk.code.len() {
        let ext = chunk.code[offset + 1];
        match Op::from_two_bytes(byte, ext) {
            Some(op) => (op, offset + 2),
            None => return (format!("UNKNOWN(0xFE 0x{:02X})", ext), offset + 2),
        }
    } else {
        match Op::from_byte(byte) {
            Some(op) => (op, offset + 1),
            None => return (format!("UNKNOWN({})", byte), offset + 1),
        }
    };
    // `base` now points past the opcode byte(s). All operand offsets below
    // used `offset + 1` for single-byte ops. Replace with `base`.
    let offset = base - 1; // compatibility: rest of code uses offset+1 for first operand

    match op {
        // No operands
        Op::drop | Op::dup |
        Op::f64_add | Op::f64_sub | Op::f64_mul | Op::f64_div | Op::f64_mod | Op::f64_neg |
        Op::f64_abs | Op::f64_ceil | Op::f64_floor | Op::f64_trunc | Op::f64_nearest | Op::f64_sqrt |
        Op::f64_min | Op::f64_max | Op::f64_copysign |
        Op::f32_abs | Op::f32_neg | Op::f32_ceil | Op::f32_floor | Op::f32_trunc | Op::f32_nearest | Op::f32_sqrt |
        Op::f32_min | Op::f32_max | Op::f32_copysign |
        Op::i32_add | Op::i32_sub | Op::i32_mul | Op::i32_div_s | Op::i32_div_u | Op::i32_rem_s | Op::i32_rem_u |
        Op::i32_rotl | Op::i32_rotr | Op::i32_clz | Op::i32_ctz | Op::i32_popcnt | Op::i32_eqz |
        Op::i64_add | Op::i64_sub | Op::i64_mul | Op::i64_div_s | Op::i64_div_u | Op::i64_rem_s | Op::i64_rem_u |
        Op::i64_and | Op::i64_or | Op::i64_xor | Op::i64_shl | Op::i64_shr_s | Op::i64_shr_u |
        Op::i64_rotl | Op::i64_rotr | Op::i64_clz | Op::i64_ctz | Op::i64_popcnt | Op::i64_eqz |
        Op::str_concat |
        Op::i32_and | Op::i32_or | Op::i32_xor | Op::i32_not | Op::i32_shl | Op::i32_shr_s | Op::i32_shr_u |
        Op::eq | Op::ne | Op::f64_lt | Op::f64_gt | Op::f64_le | Op::f64_ge |
        Op::str_lt | Op::str_gt |
        Op::bool_not |
        Op::dyn_add | Op::dyn_eq | Op::dyn_ne | Op::dyn_lt | Op::dyn_gt |
        Op::dyn_le | Op::dyn_ge | Op::dyn_neg | Op::dyn_not | Op::dyn_to_bool |
        Op::r#await | Op::set_timer |
        Op::r#return | Op::halt | Op::unreachable |
        Op::shared_new | Op::shared_array_get | Op::shared_array_set |
        Op::ref_make_weak | Op::ref_deref_weak | Op::ref_is_alive | Op::ref_register_finalizer |
        Op::memory_init | Op::memory_copy_cross |
        Op::string_as_ref | Op::string_from_ref | Op::string_ref_eq |
        Op::null | Op::undefined | Op::r#true | Op::r#false |
        Op::i32_const_0 | Op::i32_const_1 | Op::f64_const_0 |
        Op::array_get | Op::array_set |
        Op::ref_is_null | Op::ref_is_string | Op::ref_is_number | Op::ref_is_bool | Op::ref_is_object | Op::ref_is_func |
        Op::f64_from_i32 | Op::i32_from_f64 |
        Op::i32_wrap_i64 | Op::i64_extend_i32_s | Op::i64_extend_i32_u |
        Op::i64_trunc_f64_s | Op::i64_trunc_f64_u |
        Op::f64_promote_f32 | Op::f32_demote_f64 |
        Op::i32_reinterpret_f32 | Op::i64_reinterpret_f64 | Op::f32_reinterpret_i32 | Op::f64_reinterpret_i64 |
        Op::i32_extend8_s | Op::i32_extend16_s | Op::i64_extend8_s | Op::i64_extend16_s | Op::i64_extend32_s |
        Op::select |
        Op::try_end | Op::throw | Op::throw_ref |
        Op::i31_new | Op::i31_get_s | Op::i31_get_u |
        Op::inherit | Op::iter_get | Op::iter_next | Op::spread |
        Op::memory_size | Op::end | Op::unpack |
        // String builtins
        Op::str_length | Op::str_char_code_at | Op::str_from_char_code | Op::str_char_at |
        Op::str_substring | Op::str_slice | Op::str_index_of | Op::str_last_index_of |
        Op::str_equals | Op::str_compare | Op::str_to_upper | Op::str_to_lower |
        Op::str_trim | Op::str_trim_start | Op::str_trim_end |
        Op::str_starts_with | Op::str_ends_with | Op::str_contains |
        Op::str_replace | Op::str_split | Op::str_repeat |
        Op::str_pad_start | Op::str_pad_end | Op::str_reverse |
        Op::str_from_code_point | Op::str_code_point_at |
        Op::str_into_char_codes | Op::str_from_char_codes |
        Op::ref_typeof | Op::ref_is_array |
        // Array builtins
        Op::array_length | Op::array_push | Op::array_pop | Op::array_slice |
        Op::array_join | Op::array_reverse | Op::array_contains | Op::array_index_of |
        Op::array_new_default | Op::array_fill | Op::array_copy | Op::array_concat | Op::array_shift |
        // Stack switching (no-operand forms)
        Op::cont_new |
        // SIMD
        Op::i32x4_add | Op::i32x4_sub | Op::i32x4_mul | Op::i32x4_eq | Op::i32x4_gt_s | Op::i32x4_lt_s |
        Op::f64x2_add | Op::f64x2_sub | Op::f64x2_mul | Op::f64x2_div | Op::f64x2_sqrt |
        Op::f64x2_min | Op::f64x2_max | Op::f64x2_abs | Op::f64x2_neg | Op::f64x2_eq | Op::f64x2_lt | Op::f64x2_le |
        Op::f32x4_add | Op::f32x4_sub | Op::f32x4_mul | Op::f32x4_div |
        Op::i8x16_add | Op::i8x16_sub | Op::i8x16_eq |
        Op::i16x8_add | Op::i16x8_sub | Op::i16x8_mul |
        Op::v128_and | Op::v128_or | Op::v128_xor | Op::v128_not | Op::v128_andnot | Op::v128_any_true |
        Op::v128_bitselect |
        Op::i32x4_splat | Op::f64x2_splat | Op::f32x4_splat | Op::i8x16_splat | Op::i16x8_splat |
        Op::v128_load | Op::v128_store |
        // Atomics
        Op::atomic_fence |
        Op::i32_atomic_load | Op::i32_atomic_store |
        Op::i32_atomic_rmw_add | Op::i32_atomic_rmw_sub | Op::i32_atomic_rmw_and |
        Op::i32_atomic_rmw_or | Op::i32_atomic_rmw_xor | Op::i32_atomic_rmw_xchg | Op::i32_atomic_rmw_cmpxchg |
        Op::i64_atomic_load | Op::i64_atomic_store |
        Op::i64_atomic_rmw_add | Op::i64_atomic_rmw_sub | Op::i64_atomic_rmw_cmpxchg |
        Op::memory_atomic_wait32 | Op::memory_atomic_notify |
        Op::thread_spawn | Op::thread_join |
        // Memory64
        Op::i64_memory_size | Op::i64_memory_grow |
        Op::i32_load_64 | Op::i64_load_64 | Op::f64_load_64 |
        Op::i32_store_64 | Op::i64_store_64 | Op::f64_store_64 |
        // Relaxed SIMD FMA
        Op::f32x4_relaxed_madd | Op::f32x4_relaxed_nmadd |
        Op::f64x2_relaxed_madd | Op::f64x2_relaxed_nmadd |
        // Promise integration
        Op::promise_suspend => {
            (format!("{:?}", op), offset + 1)
        }

        // u16 operand
        Op::r#const | Op::local_get | Op::local_set | Op::global_get | Op::global_set |
        Op::struct_get | Op::struct_set | Op::struct_new | Op::array_new | Op::class_new | Op::method_def |
        Op::block | Op::r#loop | Op::memory_grow | Op::canon_lift | Op::canon_lower |
        Op::type_import | Op::type_export |
        Op::shared_struct_get | Op::shared_struct_set | Op::shared_struct_cas |
        Op::memory_select |
        Op::global_init |
        Op::cont_new_typed | Op::suspend_typed | Op::resume_typed |
        Op::ref_test | Op::ref_cast |
        Op::suspend | Op::resume | Op::switch => {
            let idx = chunk.read_u16(offset + 1);
            let extra = if op == Op::r#const && (idx as usize) < chunk.constants.len() {
                format!(" ({})", chunk.constants[idx as usize])
            } else {
                String::new()
            };
            (format!("{:?} {}{}", op, idx, extra), offset + 3)
        }

        // u8 operand
        Op::upvalue_get | Op::upvalue_set | Op::call | Op::call_ref | Op::str_concat_n |
        Op::return_call | Op::return_call_indirect | Op::return_call_ref |
        Op::call_indirect | Op::pack |
        Op::br_label | Op::br_if_label |
        // SIMD lane ops (u8 lane index)
        Op::i32x4_extract_lane | Op::i32x4_replace_lane | Op::i32x4_shl | Op::i32x4_shr_s | Op::i32x4_shr_u |
        Op::f64x2_extract_lane | Op::f64x2_replace_lane |
        Op::f32x4_extract_lane | Op::f32x4_replace_lane |
        Op::i8x16_extract_lane_s | Op::i8x16_extract_lane_u | Op::i8x16_replace_lane |
        Op::i16x8_extract_lane_s | Op::i16x8_extract_lane_u | Op::i16x8_replace_lane => {
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

        // br_on_cast: u16 type_name + i16 offset = 4 bytes
        Op::br_on_cast | Op::br_on_cast_fail => {
            let type_idx = chunk.read_u16(offset + 1);
            let off = chunk.read_i16(offset + 3);
            (format!("{:?} type={} offset={}", op, type_idx, off), offset + 5)
        }

        // try_table: u8 count, then count * (u8 tag + u16 offset)
        Op::try_table => {
            let count = chunk.code[offset + 1] as usize;
            let total = 2 + count * 3;
            (format!("try_table handlers={}", count), offset + total)
        }

        // v128.const: 16-byte immediate
        Op::v128_const => {
            (format!("v128.const [16 bytes]"), offset + 17)
        }

        // i8x16.shuffle: 16-byte lane indices
        Op::i8x16_shuffle => {
            (format!("i8x16.shuffle [16 indices]"), offset + 17)
        }

        // i8x16.swizzle: no operand (both vectors on stack)
        Op::i8x16_swizzle => {
            (format!("{:?}", op), offset + 1)
        }

        // Memory load/store — no operand, just opcode (addr on stack)
        Op::i32_load | Op::i32_store | Op::i64_load | Op::i64_store |
        Op::f64_load | Op::f64_store | Op::f32_load | Op::f32_store |
        Op::i32_load8_s | Op::i32_load8_u | Op::i32_load16_s | Op::i32_load16_u |
        Op::i32_store8 | Op::i32_store16 |
        Op::i64_load8_s | Op::i64_load8_u | Op::i64_load16_s | Op::i64_load16_u |
        Op::i64_load32_s | Op::i64_load32_u |
        Op::i64_store8 | Op::i64_store16 | Op::i64_store32 => {
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
        Op::set_type_id => (format!("set_type_id"), offset + 1),
    }
}
