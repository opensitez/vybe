//! Java-specific `common:java.*` dispatch.
//!
//! All patterns follow the same conventions as Go/Dart runtime adapters:
//! - Imports register on `chunks[0]`; code emits to `chunks[current]`
//! - `host::emit(chunk, module, name, argc, line)` for inline host calls
//! - `collections::emit_*(chunks, current, line)` for collection helpers
//! - `strings::emit_*(chunk, line)` for string helpers (single chunk)
//! - `core_wasm::*(&mut chunk, line, ...)` for raw WASM ops

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_compiler::compiler::instructions::{core_wasm, host};
use vybe_compiler::compiler::{collections, strings};

fn emit_stdout_text(chunk: &mut Chunk, line: u32) {
    let text_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    host::emit(chunk, "wasi:cli/stdout", "get-stdout", 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    host::emit(
        chunk,
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
        2,
        line,
    );
}

fn emit_print_stream_sentinel(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("__java_out", line);
}

fn emit_java_exp(chunks: &mut [Chunk], current: usize, upper: bool, line: u32) {
    chunks[current].emit_i32_const(6, line);
    let to_exp = chunks[current].add_import("ecma:number", "toExponential");
    chunks[current].emit_call(to_exp, 2, line);
    if upper {
        let to_upper = chunks[current].add_import("ecma:string", "toUpperCase");
        chunks[current].emit_call(to_upper, 1, line);
    }

    let (plus, plus_padded, minus, minus_padded) = if upper {
        ("E+", "E+0", "E-", "E-0")
    } else {
        ("e+", "e+0", "e-", "e-0")
    };
    chunks[current].emit_string_const(plus, line);
    chunks[current].emit_string_const(plus_padded, line);
    strings::emit_replace(&mut chunks[current], line);
    chunks[current].emit_string_const(minus, line);
    chunks[current].emit_string_const(minus_padded, line);
    strings::emit_replace(&mut chunks[current], line);
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        // ── Print ──────────────────────────────────────────────────────────
        "java.println" => {
            if argc == 0 {
                chunks[current].emit_string_const("", line);
            } else {
                let to_str = chunks[current].add_import("ecma:string", "String");
                chunks[current].emit_op_u16(Op::CALL_IMPORT, to_str, line);
                chunks[current].emit(1, line);
            }
            host::emit(&mut chunks[current], "wasi:logging/logging", "log", 1, line);
            emit_print_stream_sentinel(&mut chunks[current], line);
        }
        "java.print_no_newline" => {
            if argc == 0 {
                chunks[current].emit_string_const("", line);
            }
            // Real WASI stdout (the old target, `wasi:cli.print`, never
            // existed as a host fn). Import tables are PER CHUNK —
            // register on the chunk whose CALL_IMPORT indexes them.
            let to_str = chunks[current].add_import("ecma:string", "String");
            chunks[current].emit_op_u16(Op::CALL_IMPORT, to_str, line);
            chunks[current].emit(1, line);
            emit_stdout_text(&mut chunks[current], line);
            emit_print_stream_sentinel(&mut chunks[current], line);
        }
        "java.identity" => {}
        "java.field_set" => {
            let value = chunks[current].alloc_scratch(1);
            let field = chunks[current].alloc_scratch(1);
            let object = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
            host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
        }
        "java.field_inc" => {
            let delta = chunks[current].alloc_scratch(1);
            let field = chunks[current].alloc_scratch(1);
            let object = chunks[current].alloc_scratch(1);
            let value = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delta, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
            host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, delta, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
            host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
        }
        "java.printf" => {
            vybe_compiler::compiler::sprintf::emit_sprintf(chunks, current, argc, line);
            emit_stdout_text(&mut chunks[current], line);
            emit_print_stream_sentinel(&mut chunks[current], line);
        }
        "java.printf_array" => {
            vybe_compiler::compiler::sprintf::emit_sprintf_from_array(chunks, current, line);
            emit_stdout_text(&mut chunks[current], line);
            emit_print_stream_sentinel(&mut chunks[current], line);
        }
        "java.format_grouped_int" => {
            let to_locale = chunks[current].add_import("ecma:number", "toLocaleString");
            chunks[current].emit_call(to_locale, 1, line);
        }
        "java.format_exp_lower" => {
            emit_java_exp(chunks, current, false, line);
        }
        "java.format_exp_upper" => {
            emit_java_exp(chunks, current, true, line);
        }

        // ── Random helpers ────────────────────────────────────────────────
        "java.random_new" => {
            super::random_adapter::emit_new(chunks, current, argc, line);
        }
        "java.random_set_seed" => {
            super::random_adapter::emit_set_seed(chunks, current, line);
        }
        "java.random_next_int" => {
            super::random_adapter::emit_next_int(chunks, current, argc, line);
        }
        "java.random_next_long" => {
            super::random_adapter::emit_next_long(chunks, current, line);
        }
        "java.random_next_boolean" => {
            super::random_adapter::emit_next_bool(chunks, current, line);
        }
        "java.random_next_float" => {
            super::random_adapter::emit_next_float(chunks, current, line);
        }
        "java.random_next_double" => {
            super::random_adapter::emit_next_double(chunks, current, line);
        }
        "java.random_next_bytes" => {
            super::random_adapter::emit_next_bytes(chunks, current, line);
        }
        "java.random_split" => {
            super::random_adapter::emit_split(chunks, current, line);
        }
        "java.random_ints" => {
            super::random_adapter::emit_ints(chunks, current, argc, line);
        }
        "java.random_longs" => {
            super::random_adapter::emit_longs(chunks, current, argc, line);
        }
        "java.random_doubles" => {
            super::random_adapter::emit_doubles(chunks, current, argc, line);
        }

        // ── String helpers ─────────────────────────────────────────────────
        "java.str_is_empty" => {
            // Polymorphic: works on String, List, Map.
            core_wasm::dup(&mut chunks[current], line);
            host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
            vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            strings::emit_length(&mut chunks[current], line);
            chunks[current].emit_else(line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_end(line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
            vybe_compiler::compiler::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "java.str_is_blank" => {
            strings::emit_trim(&mut chunks[current], line);
            strings::emit_length(&mut chunks[current], line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
            vybe_compiler::compiler::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "java.is_empty" => {
            collections::emit_len(chunks, current, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
            vybe_compiler::compiler::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "java.size" => {
            super::list_adapter::emit_size(chunks, current, line);
        }
        "java.str_index_of" => {
            super::string_adapter::emit_index_of(chunks, current, argc, line);
        }
        "java.str_last_index_of" => {
            super::string_adapter::emit_last_index_of(chunks, current, argc, line);
        }
        "java.str_starts_with" => {
            super::string_adapter::emit_starts_with(chunks, current, argc, line);
        }
        "java.string_value_of" => {
            super::string_adapter::emit_value_of(chunks, current, line);
        }
        "java.string_concat" => {
            super::string_adapter::emit_concat(chunks, current, line);
        }
        "java.replace_regex" => {
            super::string_adapter::emit_replace_regex(chunks, current, true, line);
        }
        "java.replace_first_regex" => {
            super::string_adapter::emit_replace_regex(chunks, current, false, line);
        }
        "java.compare_ignore_case" => {
            super::string_adapter::emit_compare_ignore_case(chunks, current, line);
        }
        "java.equals_ignore_case" => {
            super::string_adapter::emit_equals_ignore_case(chunks, current, line);
        }
        "java.str_matches" => {
            super::string_adapter::emit_matches(chunks, current, line);
        }
        "java.to_char_array" => {
            super::string_adapter::emit_to_char_array(chunks, current, line);
        }
        "java.compare_to" => {
            super::string_adapter::emit_compare_to(chunks, current, line);
        }
        "java.char_ord" => {
            super::string_adapter::emit_char_ord(chunks, current, line);
        }

        // ── Numeric conversions ───────────────────────────────────────────
        "java.to_binary_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toBinary", 1, line);
        }
        "java.to_hex_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toHex", 1, line);
        }
        "java.to_octal_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toOctal", 1, line);
        }
        "java.parse_int" => {
            host::emit(&mut chunks[current], "ecma:number", "parseInt", argc, line);
        }

        // ── Integer bit operations (JLS java.lang.Integer) — raw WASM ops ──
        "java.int_bit_count" => {
            chunks[current].emit_op(Op::I32_POPCNT, line);
        }
        "java.int_leading_zeros" => {
            chunks[current].emit_op(Op::I32_CLZ, line);
        }
        "java.int_trailing_zeros" => {
            chunks[current].emit_op(Op::I32_CTZ, line);
        }
        "java.int_rotate_left" => {
            chunks[current].emit_op(Op::I32_ROTL, line);
        }
        "java.int_rotate_right" => {
            chunks[current].emit_op(Op::I32_ROTR, line);
        }
        "java.int_lowest_one_bit" => {
            // x & (0 - x)
            let s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            chunks[current].emit_op(Op::I32_SUB, line);
            chunks[current].emit_op(Op::I32_AND, line);
        }
        "java.int_highest_one_bit" => {
            // Smear the top bit right, then isolate it: s |= s>>>1 … s>>>16;
            // s - (s >>> 1). Branch-free, matches Integer.highestOneBit for
            // 0 and negatives (sign bit) alike.
            let s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
            for shift in [1, 2, 4, 8, 16] {
                chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
                core_wasm::i32_const(&mut chunks[current], line, shift);
                chunks[current].emit_op(Op::I32_SHR_U, line);
                chunks[current].emit_op(Op::I32_OR, line);
                chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
            }
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_SHR_U, line);
            chunks[current].emit_op(Op::I32_SUB, line);
        }
        "java.bigint_to_string" => {
            super::biginteger_adapter::emit_to_string(chunks, current, line);
        }
        "java.bigint_add" => {
            super::biginteger_adapter::emit_binary(chunks, current, "add", line);
        }
        "java.bigint_sub" => {
            super::biginteger_adapter::emit_binary(chunks, current, "sub", line);
        }
        "java.bigint_mul" => {
            super::biginteger_adapter::emit_binary(chunks, current, "mul", line);
        }
        "java.bigint_rem" => {
            super::biginteger_adapter::emit_binary(chunks, current, "rem", line);
        }
        "java.bigint_pow" => {
            super::biginteger_adapter::emit_binary(chunks, current, "pow", line);
        }
        "java.bigint_and" => {
            super::biginteger_adapter::emit_binary(chunks, current, "and", line);
        }
        "java.bigint_or" => {
            super::biginteger_adapter::emit_binary(chunks, current, "or", line);
        }
        "java.bigint_xor" => {
            super::biginteger_adapter::emit_binary(chunks, current, "xor", line);
        }
        "java.bigint_shl" => {
            super::biginteger_adapter::emit_binary(chunks, current, "shl", line);
        }
        "java.bigint_shr" => {
            super::biginteger_adapter::emit_binary(chunks, current, "shr", line);
        }
        "java.bigint_neg" => {
            super::biginteger_adapter::emit_unary(chunks, current, "neg", line);
        }
        "java.bigint_not" => {
            super::biginteger_adapter::emit_unary(chunks, current, "not", line);
        }
        "java.bigint_abs" => {
            super::biginteger_adapter::emit_abs(chunks, current, line);
        }
        "java.bigint_compare_to" => {
            super::biginteger_adapter::emit_compare_to(chunks, current, line);
        }
        "java.bigint_signum" => {
            super::biginteger_adapter::emit_signum(chunks, current, line);
        }
        "java.bigint_max" => {
            super::biginteger_adapter::emit_min_max(chunks, current, false, line);
        }
        "java.bigint_min" => {
            super::biginteger_adapter::emit_min_max(chunks, current, true, line);
        }
        "java.bigint_bit_length" => {
            super::biginteger_adapter::emit_bit_length(chunks, current, line);
        }
        "java.bigint_test_bit" => {
            super::biginteger_adapter::emit_test_bit(chunks, current, line);
        }
        "java.bigint_gcd" => {
            super::biginteger_adapter::emit_gcd(chunks, current, line);
        }
        "java.bigint_is_probable_prime" => {
            super::biginteger_adapter::emit_is_probable_prime(chunks, current, line);
        }
        "java.bigint_next_probable_prime" => {
            super::biginteger_adapter::emit_next_probable_prime(chunks, current, line);
        }
        "java.is_infinite" => {
            host::emit(&mut chunks[current], "ecma:number", "isFinite", 1, line);
            vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
        }
        "java.signum" => {
            host::emit(&mut chunks[current], "ecma:math", "sign", 1, line);
        }
        "java.math_scalb" => {
            super::math_adapter::emit_scalb(chunks, current, line);
        }
        "java.math_ulp" => {
            super::math_adapter::emit_ulp(chunks, current, line);
        }
        "java.math_get_exponent" => {
            super::math_adapter::emit_get_exponent(chunks, current, line);
        }
        "java.math_copy_sign" => {
            super::math_adapter::emit_copy_sign(chunks, current, line);
        }
        "java.math_next_after" => {
            super::math_adapter::emit_next_after(chunks, current, line);
        }
        "java.math_next_up" => {
            super::math_adapter::emit_next_up(chunks, current, line);
        }
        "java.math_next_down" => {
            super::math_adapter::emit_next_down(chunks, current, line);
        }
        "java.math_fma" => {
            super::math_adapter::emit_fma(chunks, current, line);
        }
        "java.math_expm1" => {
            super::math_adapter::emit_expm1(chunks, current, line);
        }
        "java.math_log1p" => {
            super::math_adapter::emit_log1p(chunks, current, line);
        }
        "java.math_to_degrees" => {
            super::math_adapter::emit_to_degrees(chunks, current, line);
        }
        "java.math_to_radians" => {
            super::math_adapter::emit_to_radians(chunks, current, line);
        }
        "java.math_ieee_remainder" => {
            super::math_adapter::emit_ieee_remainder(chunks, current, line);
        }
        "java.math_add_exact" => {
            super::math_adapter::emit_add_exact(chunks, current, line);
        }
        "java.math_subtract_exact" => {
            super::math_adapter::emit_subtract_exact(chunks, current, line);
        }
        "java.math_multiply_exact" => {
            super::math_adapter::emit_multiply_exact(chunks, current, line);
        }
        "java.math_increment_exact" => {
            super::math_adapter::emit_increment_exact(chunks, current, line);
        }
        "java.math_decrement_exact" => {
            super::math_adapter::emit_decrement_exact(chunks, current, line);
        }
        "java.math_negate_exact" => {
            super::math_adapter::emit_negate_exact(chunks, current, line);
        }
        "java.floor_div" => {
            super::math_adapter::emit_floor_div(chunks, current, line);
        }
        "java.floor_mod" => {
            super::math_adapter::emit_floor_mod(chunks, current, line);
        }
        "java.compare" => {
            let b_slot = chunks[current].alloc_scratch(1);
            let a_slot = chunks[current].alloc_scratch(1);
            let result_slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
            vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
            vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(-1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
            vybe_compiler::compiler::ops::emit_dyn_gt(&mut chunks[current], line);
            vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
        }
        "java.trunc_cast" => {
            super::string_adapter::emit_trunc_cast(chunks, current, line);
        }
        "java.double_to_string" => {
            super::list_adapter::emit_double_to_string(chunks, current, line);
        }
        "java.uuid_from_string" => {
            super::uuid_adapter::emit_from_string(chunks, current, line);
        }
        "java.uuid_name_from_bytes" => {
            super::uuid_adapter::emit_name_from_bytes(chunks, current, line);
        }
        "java.uuid_version" => {
            super::uuid_adapter::emit_version(chunks, current, line);
        }
        "java.uuid_variant" => {
            super::uuid_adapter::emit_variant(chunks, current, line);
        }
        "java.uuid_most_bits" => {
            super::uuid_adapter::emit_most_bits(chunks, current, line);
        }
        "java.uuid_least_bits" => {
            super::uuid_adapter::emit_least_bits(chunks, current, line);
        }
        "java.uuid_compare_to" => {
            super::uuid_adapter::emit_compare_to(chunks, current, line);
        }
        "java.uuid_hash_code" => {
            super::uuid_adapter::emit_hash_code(chunks, current, line);
        }
        "java.uuid_new" => {
            super::uuid_adapter::emit_new(chunks, current, argc, line);
        }
        "java.bitset_new" => {
            super::bitset_adapter::emit_new(chunks, current, argc, line);
        }
        "java.bitset_value_of" => {
            super::bitset_adapter::emit_value_of(chunks, current, line);
        }
        "java.bitset_set" => {
            super::bitset_adapter::emit_set(chunks, current, argc, line);
        }
        "java.bitset_get" => {
            super::bitset_adapter::emit_get(chunks, current, argc, line);
        }
        "java.bitset_clear" => {
            super::bitset_adapter::emit_clear(chunks, current, argc, line);
        }
        "java.bitset_flip" => {
            super::bitset_adapter::emit_flip(chunks, current, argc, line);
        }
        "java.bitset_cardinality" => {
            super::bitset_adapter::emit_cardinality(chunks, current, line);
        }
        "java.bitset_length" => {
            super::bitset_adapter::emit_length(chunks, current, line);
        }
        "java.bitset_size" => {
            super::bitset_adapter::emit_size(chunks, current, line);
        }
        "java.bitset_is_empty" => {
            super::bitset_adapter::emit_is_empty(chunks, current, line);
        }
        "java.bitset_next_set_bit" => {
            super::bitset_adapter::emit_next_set_bit(chunks, current, line);
        }
        "java.bitset_next_clear_bit" => {
            super::bitset_adapter::emit_next_clear_bit(chunks, current, line);
        }
        "java.bitset_previous_set_bit" => {
            super::bitset_adapter::emit_previous_set_bit(chunks, current, line);
        }
        "java.bitset_previous_clear_bit" => {
            super::bitset_adapter::emit_previous_clear_bit(chunks, current, line);
        }
        "java.bitset_and" => {
            super::bitset_adapter::emit_and(chunks, current, line);
        }
        "java.bitset_or" => {
            super::bitset_adapter::emit_or(chunks, current, line);
        }
        "java.bitset_xor" => {
            super::bitset_adapter::emit_xor(chunks, current, line);
        }
        "java.bitset_and_not" => {
            super::bitset_adapter::emit_and_not(chunks, current, line);
        }
        "java.bitset_intersects" => {
            super::bitset_adapter::emit_intersects(chunks, current, line);
        }
        "java.bitset_equals" => {
            super::bitset_adapter::emit_equals(chunks, current, line);
        }
        "java.bitset_clone" => {
            super::bitset_adapter::emit_clone(chunks, current, line);
        }
        "java.bitset_stream" => {
            super::bitset_adapter::emit_stream(chunks, current, line);
        }
        "java.bitset_to_array" => {
            super::bitset_adapter::emit_to_array(chunks, current, line);
        }
        "java.bitset_to_string" => {
            super::bitset_adapter::emit_to_string(chunks, current, line);
        }
        "java.bitset_hash_code" => {
            super::bitset_adapter::emit_hash_code(chunks, current, line);
        }
        "java.enum_set_none_of" => {
            super::enum_set_adapter::emit_none_of(chunks, current, line);
        }
        "java.enum_set_all_of" => {
            super::enum_set_adapter::emit_all_of(chunks, current, line);
        }
        "java.enum_set_of" => {
            super::enum_set_adapter::emit_of(chunks, current, argc, line);
        }
        "java.enum_set_copy_of" => {
            super::enum_set_adapter::emit_copy_of(chunks, current, line);
        }
        "java.enum_set_complement_of" => {
            super::enum_set_adapter::emit_complement_of(chunks, current, line);
        }
        "java.enum_set_range" => {
            super::enum_set_adapter::emit_range(chunks, current, line);
        }
        "java.enum_set_add" => {
            super::enum_set_adapter::emit_add(chunks, current, line);
        }
        "java.enum_set_add_all" => {
            super::enum_set_adapter::emit_add_all(chunks, current, line);
        }
        "java.enum_set_contains" => {
            super::enum_set_adapter::emit_contains(chunks, current, line);
        }
        "java.enum_set_contains_all" => {
            super::enum_set_adapter::emit_contains_all(chunks, current, line);
        }
        "java.enum_set_remove" => {
            super::enum_set_adapter::emit_remove(chunks, current, line);
        }
        "java.enum_set_equals" => {
            super::enum_set_adapter::emit_equals(chunks, current, line);
        }
        "java.enum_set_hash_code" => {
            super::enum_set_adapter::emit_hash_code(chunks, current, line);
        }
        "java.enum_set_iterator" => {
            super::enum_set_adapter::emit_iterator(chunks, current, line);
        }
        "java.enum_set_get_class" => {
            super::enum_set_adapter::emit_get_class(chunks, current, line);
        }
        "java.instant_of_epoch_second" => {
            super::instant_adapter::emit_of_epoch_second(chunks, current, argc, line);
        }
        "java.instant_of_epoch_milli" => {
            super::instant_adapter::emit_of_epoch_milli(chunks, current, line);
        }
        "java.instant_parse" => {
            super::instant_adapter::emit_parse(chunks, current, line);
        }
        "java.local_date_of" => {
            super::instant_adapter::emit_local_date_of(chunks, current, line);
        }
        "java.local_date_parse" => {
            super::instant_adapter::emit_local_date_parse(chunks, current, line);
        }
        "java.local_time_of" => {
            super::instant_adapter::emit_local_time_of(chunks, current, argc, line);
        }
        "java.local_time_parse" => {
            super::instant_adapter::emit_local_time_parse(chunks, current, line);
        }
        "java.local_datetime_of" => {
            super::instant_adapter::emit_local_datetime_of(chunks, current, argc, line);
        }
        "java.local_datetime_parse" => {
            super::instant_adapter::emit_local_datetime_parse(chunks, current, line);
        }
        "java.offset_datetime_of" => {
            super::instant_adapter::emit_offset_datetime_of(chunks, current, argc, line);
        }
        "java.offset_datetime_of_instant" => {
            super::instant_adapter::emit_offset_datetime_of_instant(chunks, current, line);
        }
        "java.offset_datetime_parse" => {
            super::instant_adapter::emit_offset_datetime_parse(chunks, current, line);
        }
        "java.zoned_datetime_of" => {
            super::instant_adapter::emit_zoned_datetime_of(chunks, current, argc, line);
        }
        "java.zoned_datetime_of_instant" => {
            super::instant_adapter::emit_zoned_datetime_of_instant(chunks, current, line);
        }
        "java.zoned_datetime_of_strict" => {
            super::instant_adapter::emit_zoned_datetime_of_strict(chunks, current, line);
        }
        "java.zoned_datetime_parse" => {
            super::instant_adapter::emit_zoned_datetime_parse(chunks, current, line);
        }
        "java.instant_get_epoch_second" => {
            super::instant_adapter::emit_get_epoch_second(chunks, current, line);
        }
        "java.instant_get_nano" => {
            super::instant_adapter::emit_get_nano(chunks, current, line);
        }
        "java.instant_to_epoch_milli" => {
            super::instant_adapter::emit_to_epoch_milli(chunks, current, line);
        }
        "java.instant_plus_seconds" => {
            super::instant_adapter::emit_plus_seconds(chunks, current, 1.0, line);
        }
        "java.instant_minus_seconds" => {
            super::instant_adapter::emit_plus_seconds(chunks, current, -1.0, line);
        }
        "java.instant_plus_millis" => {
            super::instant_adapter::emit_plus_millis(chunks, current, 1.0, line);
        }
        "java.instant_minus_millis" => {
            super::instant_adapter::emit_plus_millis(chunks, current, -1.0, line);
        }
        "java.instant_plus_nanos" => {
            super::instant_adapter::emit_plus_nanos(chunks, current, 1.0, line);
        }
        "java.instant_minus_nanos" => {
            super::instant_adapter::emit_plus_nanos(chunks, current, -1.0, line);
        }
        "java.instant_compare_to" => {
            super::instant_adapter::emit_compare(chunks, current, line);
        }
        "java.instant_is_before" => {
            super::instant_adapter::emit_is_before_after(chunks, current, false, line);
        }
        "java.instant_is_after" => {
            super::instant_adapter::emit_is_before_after(chunks, current, true, line);
        }
        "java.instant_equals" => {
            super::instant_adapter::emit_equals(chunks, current, line);
        }
        "java.instant_to_string" => {
            super::instant_adapter::emit_to_string(chunks, current, line);
        }
        "java.duration_of_hours" => {
            super::instant_adapter::emit_duration_hours(chunks, current, line);
        }
        "java.duration_of_minutes" => {
            super::instant_adapter::emit_duration_minutes(chunks, current, line);
        }
        "java.duration_of_seconds" => {
            super::instant_adapter::emit_duration_seconds(chunks, current, line);
        }
        "java.duration_between" => {
            super::instant_adapter::emit_duration_between(chunks, current, line);
        }
        "java.period_days" => {
            super::instant_adapter::emit_period_days(chunks, current, line);
        }
        "java.period_months" => {
            super::instant_adapter::emit_period_months(chunks, current, line);
        }
        "java.period_between" => {
            super::instant_adapter::emit_period_between(chunks, current, line);
        }
        "java.class_is_instance" => {
            super::class_adapter::emit_is_instance(chunks, current, line);
        }
        "java.class_name" => {
            super::reflection_adapter::emit_class_name(chunks, current, line);
        }
        "java.class_simple_name" => {
            super::reflection_adapter::emit_class_simple_name(chunks, current, line);
        }
        "java.object_get_class" => {
            super::reflection_adapter::emit_object_get_class(chunks, current, line);
        }
        "java.zone_offset_of_hours" => {
            super::instant_adapter::emit_zone_offset_hours(chunks, current, line);
        }
        "java.zone_id_of" => {
            super::instant_adapter::emit_zone_id_utc(chunks, current, line);
        }
        "java.zone_id_system_default" => {
            super::instant_adapter::emit_zone_id_system_default(chunks, current, line);
        }
        "java.zone_id_short_ids" => {
            super::instant_adapter::emit_zone_id_short_ids(chunks, current, line);
        }
        "java.zone_id_from" => {
            super::instant_adapter::emit_zone_id_from(chunks, current, line);
        }
        "java.zone_id_of_offset" => {
            super::instant_adapter::emit_zone_id_of_offset(chunks, current, line);
        }
        "java.zone_normalized" => {
            super::instant_adapter::emit_zone_normalized(chunks, current, line);
        }
        "java.zone_display_name" => {
            super::instant_adapter::emit_zone_display_name(chunks, current, argc, line);
        }
        "java.zone_rules_fixed" => {
            super::instant_adapter::emit_zone_rules_fixed(chunks, current, line);
        }
        "java.zone_rules_get_offset" => {
            super::instant_adapter::emit_zone_rules_get_offset(chunks, current, line);
        }
        "java.zone_offset_total_seconds" => {
            super::instant_adapter::emit_get_total_seconds(chunks, current, line);
        }
        "java.zone_compare_to" => {
            super::instant_adapter::emit_zone_compare_to(chunks, current, line);
        }
        "java.zone_hash_code" => {
            super::instant_adapter::emit_zone_hash_code(chunks, current, line);
        }
        "java.instant_with_offset" => {
            super::instant_adapter::emit_with_offset(chunks, current, line);
        }
        "java.instant_with_zone" => {
            super::instant_adapter::emit_with_zone_same_instant(chunks, current, line);
        }
        "java.instant_get_offset" => {
            super::instant_adapter::emit_get_offset(chunks, current, line);
        }
        "java.instant_get_zone" => {
            super::instant_adapter::emit_get_zone(chunks, current, line);
        }
        "java.instant_get_year" => {
            super::instant_adapter::emit_component(chunks, current, "getUTCFullYear", false, line);
        }
        "java.instant_get_month" => {
            super::instant_adapter::emit_component(chunks, current, "getUTCMonth", true, line);
        }
        "java.instant_get_day" => {
            super::instant_adapter::emit_component(chunks, current, "getUTCDate", false, line);
        }
        "java.instant_get_hour" => {
            super::instant_adapter::emit_component(chunks, current, "getUTCHours", false, line);
        }
        "java.instant_get_minute" => {
            super::instant_adapter::emit_component(chunks, current, "getUTCMinutes", false, line);
        }
        "java.instant_get_second" => {
            super::instant_adapter::emit_component(chunks, current, "getUTCSeconds", false, line);
        }
        "java.instant_to_local_date" => {
            super::instant_adapter::emit_local_date_string(chunks, current, line);
        }
        "java.time_to_string" => {
            super::instant_adapter::emit_time_to_string(chunks, current, line);
        }
        "java.time_format" => {
            super::instant_adapter::emit_time_format(chunks, current, line);
        }
        "java.time_plus_days" => {
            super::instant_adapter::emit_time_plus_unit(chunks, current, 1.0, 86400.0, line);
        }
        "java.time_minus_days" => {
            super::instant_adapter::emit_time_plus_unit(chunks, current, -1.0, 86400.0, line);
        }
        "java.time_plus_weeks" => {
            super::instant_adapter::emit_time_plus_unit(chunks, current, 1.0, 604800.0, line);
        }
        "java.time_plus_months" => {
            super::instant_adapter::emit_time_plus_months(chunks, current, 1.0, line);
        }
        "java.time_minus_months" => {
            super::instant_adapter::emit_time_plus_months(chunks, current, -1.0, line);
        }
        "java.time_plus_hours" => {
            super::instant_adapter::emit_time_plus_unit(chunks, current, 1.0, 3600.0, line);
        }
        "java.time_minus_hours" => {
            super::instant_adapter::emit_time_plus_unit(chunks, current, -1.0, 3600.0, line);
        }
        "java.time_plus_minutes" => {
            super::instant_adapter::emit_time_plus_unit(chunks, current, 1.0, 60.0, line);
        }
        "java.time_minus_minutes" => {
            super::instant_adapter::emit_time_plus_unit(chunks, current, -1.0, 60.0, line);
        }
        "java.time_with_year" => {
            super::instant_adapter::emit_time_with_field(
                chunks,
                current,
                "setUTCFullYear",
                false,
                line,
            );
        }
        "java.time_with_month" => {
            super::instant_adapter::emit_time_with_field(
                chunks,
                current,
                "setUTCMonth",
                true,
                line,
            );
        }
        "java.time_with_day" => {
            super::instant_adapter::emit_time_with_field(
                chunks,
                current,
                "setUTCDate",
                false,
                line,
            );
        }
        "java.time_with_hour" => {
            super::instant_adapter::emit_time_with_field(
                chunks,
                current,
                "setUTCHours",
                false,
                line,
            );
        }
        "java.time_with_minute" => {
            super::instant_adapter::emit_time_with_field(
                chunks,
                current,
                "setUTCMinutes",
                false,
                line,
            );
        }
        "java.time_with_second" => {
            super::instant_adapter::emit_time_with_field(
                chunks,
                current,
                "setUTCSeconds",
                false,
                line,
            );
        }
        "java.time_length_of_month" => {
            super::instant_adapter::emit_time_length_of_month(chunks, current, line);
        }
        "java.time_range_day" => {
            super::instant_adapter::emit_time_range_day(chunks, current, line);
        }
        "java.time_is_leap_year" => {
            super::instant_adapter::emit_time_is_leap_year(chunks, current, line);
        }
        "java.time_day_of_year" => {
            super::instant_adapter::emit_time_day_of_year(chunks, current, line);
        }
        "java.time_day_of_week" => {
            super::instant_adapter::emit_time_day_of_week(chunks, current, line);
        }
        "java.duration_to_hours" => {
            super::instant_adapter::emit_duration_to_hours(chunks, current, line);
        }
        "java.duration_to_minutes" => {
            super::instant_adapter::emit_duration_to_minutes(chunks, current, line);
        }
        "java.duration_plus_hours" => {
            super::instant_adapter::emit_duration_plus_hours(chunks, current, 1.0, line);
        }
        "java.duration_minus_minutes" => {
            super::instant_adapter::emit_duration_plus_minutes(chunks, current, -1.0, line);
        }
        "java.time_with_offset_same_local" => {
            super::instant_adapter::emit_time_with_offset_same_local(chunks, current, line);
        }
        "java.time_with_zone_same_local" => {
            super::instant_adapter::emit_with_zone_same_local(chunks, current, line);
        }
        "java.zoned_later_overlap" => {
            super::instant_adapter::emit_overlap_offset(chunks, current, 1, line);
        }
        "java.zoned_earlier_overlap" => {
            super::instant_adapter::emit_overlap_offset(chunks, current, 2, line);
        }
        "java.instant_truncated" => {
            super::instant_adapter::emit_truncated(chunks, current, line);
        }
        "java.instant_hash_code" => {
            super::instant_adapter::emit_hash_code(chunks, current, line);
        }

        // ── Character helpers ─────────────────────────────────────────────
        "java.char_is_digit" => {
            host::emit(&mut chunks[current], "ecma:char", "isDigit", 1, line);
        }
        "java.char_is_letter" => {
            host::emit(&mut chunks[current], "ecma:char", "isLetter", 1, line);
        }
        "java.char_is_alnum" => {
            host::emit(&mut chunks[current], "ecma:char", "isAlnum", 1, line);
        }
        "java.char_is_upper" => {
            host::emit(&mut chunks[current], "ecma:char", "isUpper", 1, line);
        }
        "java.char_is_lower" => {
            host::emit(&mut chunks[current], "ecma:char", "isLower", 1, line);
        }
        "java.char_is_space" => {
            host::emit(&mut chunks[current], "ecma:char", "isSpace", 1, line);
        }
        "java.char_to_upper" => {
            strings::emit_to_upper(&mut chunks[current], line);
        }
        "java.char_to_lower" => {
            strings::emit_to_lower(&mut chunks[current], line);
        }
        "java.char_numeric" => {
            host::emit(&mut chunks[current], "ecma:number", "parseInt", 1, line);
        }

        // ── Array helpers ─────────────────────────────────────────────────
        "java.new_array" => {
            collections::emit_new_with_length(chunks, current, line);
        }
        "java.new_int_array" => {
            emit_new_array_with_default(chunks, current, line, JavaArrayDefault::IntZero);
        }
        "java.new_bool_array" => {
            emit_new_array_with_default(chunks, current, line, JavaArrayDefault::BoolFalse);
        }
        "java.new_int_2d_array" => {
            super::arrays_adapter::emit_new_int_2d(chunks, current, line);
        }
        "java.array_clone" => {
            collections::emit_slice(chunks, current, line);
        }
        "java.arrays_sort" => {
            super::arrays_adapter::emit_sort(chunks, current, argc, line);
        }
        "java.arrays_fill" => {
            super::arrays_adapter::emit_fill(chunks, current, argc, line);
        }
        "java.arrays_copy_of" => {
            super::arrays_adapter::emit_copy_of(chunks, current, line);
        }
        "java.arrays_copy_of_range" => {
            super::arrays_adapter::emit_copy_of_range(chunks, current, line);
        }
        "java.arrays_to_string" => {
            super::arrays_adapter::emit_to_string(chunks, current, line);
        }
        "java.arrays_deep_to_string" => {
            super::arrays_adapter::emit_deep_to_string(chunks, current, line);
        }
        "java.arrays_equals" => {
            super::arrays_adapter::emit_equals(chunks, current, line);
        }
        "java.arrays_deep_equals" => {
            super::arrays_adapter::emit_deep_equals(chunks, current, line);
        }
        "java.arrays_compare" => {
            super::arrays_adapter::emit_compare(chunks, current, line);
        }
        "java.arrays_compare_unsigned" => {
            super::arrays_adapter::emit_compare_unsigned(chunks, current, line);
        }
        "java.arrays_mismatch" => {
            super::arrays_adapter::emit_mismatch(chunks, current, line);
        }
        "java.arrays_set_all" => {
            super::arrays_adapter::emit_set_all(chunks, current, line);
        }
        "java.arrays_parallel_prefix" => {
            super::arrays_adapter::emit_parallel_prefix(chunks, current, argc, line);
        }
        "java.arrays_binary_search" => {
            super::arrays_adapter::emit_binary_search(chunks, current, line);
        }
        "java.arrays_as_list" => {
            super::list_adapter::emit_arrays_as_list(chunks, current, argc, line);
        }

        // ── Primitive stream helpers ─────────────────────────────────────
        "java.stream_empty" => {
            super::stream_adapter::emit_empty(chunks, current, line);
        }
        "java.stream_of" => {
            super::stream_adapter::emit_of(chunks, current, argc, line);
        }
        "java.stream_builder" => {
            super::stream_adapter::emit_builder(chunks, current, line);
        }
        "java.stream_builder_add" => {
            super::stream_adapter::emit_builder_add(chunks, current, line);
        }
        "java.stream_range" => {
            super::stream_adapter::emit_range(chunks, current, false, line);
        }
        "java.stream_range_closed" => {
            super::stream_adapter::emit_range(chunks, current, true, line);
        }
        "java.stream_concat" => {
            super::stream_adapter::emit_concat(chunks, current, line);
        }
        "java.collectors_joining" => {
            super::stream_adapter::emit_collectors_joining(chunks, current, argc, line);
        }
        "java.collectors_to_list" => {
            super::stream_adapter::emit_collectors_to_list(chunks, current, line);
        }
        "java.collectors_to_set" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "toSet", 0, line);
        }
        "java.collectors_to_collection" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "toCollection", 1, line);
        }
        "java.collectors_counting" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "counting", 0, line);
        }
        "java.collectors_summing_int" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "summingInt", 1, line);
        }
        "java.collectors_averaging_int" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "averagingInt", 1, line);
        }
        "java.collectors_to_map" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "toMap", 2, line);
        }
        "java.collectors_mapping" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "mapping", 2, line);
        }
        "java.collectors_filtering" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "filtering", 2, line);
        }
        "java.collectors_collecting_and_then" => {
            super::stream_adapter::emit_collector_tag(
                chunks,
                current,
                "collectingAndThen",
                2,
                line,
            );
        }
        "java.collectors_reducing" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "reducing", argc, line);
        }
        "java.collectors_min_by" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "minBy", 1, line);
        }
        "java.collectors_max_by" => {
            super::stream_adapter::emit_collector_tag(chunks, current, "maxBy", 1, line);
        }
        "java.collectors_grouping_by" => {
            super::stream_adapter::emit_collector_tag_with_default_downstream(
                chunks,
                current,
                "groupingBy",
                argc,
                line,
            );
        }
        "java.collectors_partitioning_by" => {
            super::stream_adapter::emit_collector_tag_with_default_downstream(
                chunks,
                current,
                "partitioningBy",
                argc,
                line,
            );
        }
        "java.stream_collect" => {
            super::stream_adapter::emit_collect(chunks, current, line);
        }
        "java.stream_generate" => {
            super::stream_adapter::emit_generate(chunks, current, line);
        }
        "java.stream_iterate" => {
            super::stream_adapter::emit_iterate(chunks, current, argc, line);
        }
        "java.stream_iterate_strict" => {
            super::stream_adapter::emit_iterate_strict(chunks, current, argc, line);
        }
        "java.stream_count" => {
            super::stream_adapter::emit_count(chunks, current, line);
        }
        "java.stream_to_array" => {
            super::stream_adapter::emit_to_array(chunks, current, argc, line);
        }
        "java.stream_sum" => {
            super::stream_adapter::emit_sum(chunks, current, line);
        }
        "java.stream_map" => {
            super::stream_adapter::emit_map(chunks, current, line);
        }
        "java.stream_filter" => {
            super::stream_adapter::emit_filter(chunks, current, line);
        }
        "java.stream_peek" => {
            super::stream_adapter::emit_peek(chunks, current, line);
        }
        "java.stream_distinct" => {
            super::stream_adapter::emit_distinct(chunks, current, line);
        }
        "java.stream_flat_map" => {
            super::stream_adapter::emit_flat_map(chunks, current, line);
        }
        "java.stream_sorted" => {
            super::stream_adapter::emit_sorted(chunks, current, argc, line);
        }
        "java.stream_limit" => {
            super::stream_adapter::emit_limit(chunks, current, line);
        }
        "java.stream_skip" => {
            super::stream_adapter::emit_skip(chunks, current, line);
        }
        "java.stream_take_while" => {
            super::stream_adapter::emit_take_while(chunks, current, line);
        }
        "java.stream_drop_while" => {
            super::stream_adapter::emit_drop_while(chunks, current, line);
        }
        "java.stream_find_first" => {
            super::stream_adapter::emit_find_first(chunks, current, line);
        }
        "java.stream_min" => {
            super::stream_adapter::emit_extreme_value(chunks, current, argc, true, line);
        }
        "java.stream_max" => {
            super::stream_adapter::emit_extreme_value(chunks, current, argc, false, line);
        }
        "java.stream_max_value" => {
            super::stream_adapter::emit_max_value(chunks, current, line);
        }
        "java.stream_average" => {
            super::stream_adapter::emit_average(chunks, current, line);
        }
        "java.stream_average_value" => {
            super::stream_adapter::emit_average_value(chunks, current, line);
        }
        "java.stream_any_match" => {
            super::stream_adapter::emit_any_match(chunks, current, line);
        }
        "java.stream_all_match" => {
            super::stream_adapter::emit_all_match(chunks, current, line);
        }
        "java.stream_none_match" => {
            super::stream_adapter::emit_none_match(chunks, current, line);
        }
        "java.stream_reduce" => {
            super::stream_adapter::emit_reduce(chunks, current, argc, line);
        }
        "java.stream_for_each" => {
            super::stream_adapter::emit_for_each(chunks, current, line);
        }
        "java.stream_optional_get" => {
            super::stream_adapter::emit_get_optional_value(chunks, current, line);
        }

        // ── List helpers ──────────────────────────────────────────────────
        "java.mutable_list_of" => {
            collections::emit_array_new(chunks, current, argc as u16, line);
        }
        "java.copy_on_write_list_new" => {
            super::list_adapter::emit_copy_on_write_list_new(chunks, current, argc, line);
        }
        "java.linked_blocking_queue_new" => {
            super::list_adapter::emit_linked_blocking_queue_new(chunks, current, argc, line);
        }
        "java.vector_new" => {
            super::list_adapter::emit_vector_new(chunks, current, argc, line);
        }
        "java.hash_set_new" => {
            collections::emit_array_new(chunks, current, argc as u16, line);
            super::list_adapter::emit_mark_set_collection(chunks, current, line);
        }
        "java.list_of" => {
            collections::emit_array_new(chunks, current, argc as u16, line);
            super::list_adapter::emit_mark_immutable_list(chunks, current, line);
        }
        "java.set_of" => {
            super::list_adapter::emit_set_of(chunks, current, argc, line);
        }
        "java.set_copy_of" => {
            collections::emit_clone(chunks, current, line);
            super::list_adapter::emit_mark_set_collection(chunks, current, line);
            super::list_adapter::emit_mark_immutable_list(chunks, current, line);
        }
        "java.list_copy_of" => {
            collections::emit_clone(chunks, current, line);
            super::list_adapter::emit_mark_immutable_list(chunks, current, line);
        }
        "java.map_of" => {
            super::list_adapter::emit_map_of(chunks, current, argc, line);
        }
        "java.map_entry" => {
            super::list_adapter::emit_map_entry(chunks, current, line);
        }
        "java.map_of_entries" => {
            super::list_adapter::emit_map_of_entries(chunks, current, argc, line);
        }
        "java.empty_list" => {
            collections::emit_array_new(chunks, current, 0, line);
            super::list_adapter::emit_mark_immutable_list(chunks, current, line);
        }
        "java.singleton_list" => {
            host::emit(&mut chunks[current], "ecma:array", "of", 1, line);
            super::list_adapter::emit_mark_immutable_list(chunks, current, line);
        }
        "java.n_copies" => {
            super::list_adapter::emit_n_copies(chunks, current, line);
        }
        "java.list_get" => {
            collections::emit_get(chunks, current, line);
        }
        "java.list_clone" => {
            super::list_adapter::emit_list_clone(chunks, current, line);
        }
        "java.list_index_of" => {
            collections::emit_index_of(chunks, current, line);
        }
        "java.add" => {
            super::list_adapter::emit_add(chunks, current, argc, line);
        }
        "java.copy_on_write_add_if_absent" => {
            super::list_adapter::emit_copy_on_write_add_if_absent(chunks, current, line);
        }
        "java.iterator_remove_unsupported" => {
            super::list_adapter::emit_iterator_remove_unsupported(chunks, current, line);
        }
        "java.blocking_queue_add" => {
            super::list_adapter::emit_blocking_queue_offer(chunks, current, argc, true, line);
        }
        "java.blocking_queue_offer" => {
            super::list_adapter::emit_blocking_queue_offer(chunks, current, argc, false, line);
        }
        "java.blocking_queue_put" => {
            super::list_adapter::emit_blocking_queue_put(chunks, current, line);
        }
        "java.blocking_queue_take" => {
            super::list_adapter::emit_blocking_queue_take(chunks, current, line);
        }
        "java.blocking_queue_poll" => {
            super::list_adapter::emit_blocking_queue_poll(chunks, current, argc, line);
        }
        "java.queue_remove_checked" => {
            super::list_adapter::emit_queue_remove_checked(chunks, current, line);
        }
        "java.queue_element_checked" => {
            super::list_adapter::emit_queue_element_checked(chunks, current, line);
        }
        "java.blocking_queue_remaining_capacity" => {
            super::list_adapter::emit_blocking_queue_remaining_capacity(chunks, current, line);
        }
        "java.blocking_queue_drain_to" => {
            super::list_adapter::emit_blocking_queue_drain_to(chunks, current, argc, line);
        }
        "java.sorted_add" => {
            super::list_adapter::emit_sorted_add(chunks, current, line);
        }
        "java.priority_add" => {
            super::list_adapter::emit_priority_add(chunks, current, line);
        }
        "java.sorted_poll" => {
            super::list_adapter::emit_queue_poll(chunks, current, line);
        }
        "java.map_get" => {
            super::list_adapter::emit_map_get(chunks, current, line);
        }
        "java.sorted_set_new" => {
            super::list_adapter::emit_sorted_collection_new(chunks, current, argc, false, line);
        }
        "java.priority_queue_new" => {
            super::list_adapter::emit_priority_queue_new(chunks, current, argc, line);
        }
        "java.sorted_map_new" => {
            super::list_adapter::emit_sorted_collection_new(chunks, current, argc, true, line);
        }
        "java.hash_map_new" => {
            super::list_adapter::emit_hash_map_new(chunks, current, argc, line);
        }
        "java.concurrent_hash_map_new" => {
            super::list_adapter::emit_concurrent_hash_map_new(chunks, current, argc, line);
        }
        "java.identity_hash_map_new" => {
            super::list_adapter::emit_identity_hash_map_new(chunks, current, argc, line);
        }
        "java.linked_hash_map_new" => {
            super::list_adapter::emit_linked_hash_map_new(chunks, current, argc, line);
        }
        "java.concurrent_for_each_key" => {
            super::list_adapter::emit_concurrent_for_each(chunks, current, 0, line);
        }
        "java.concurrent_for_each_value" => {
            super::list_adapter::emit_concurrent_for_each(chunks, current, 1, line);
        }
        "java.concurrent_reduce_keys" => {
            super::list_adapter::emit_concurrent_reduce(chunks, current, 0, line);
        }
        "java.concurrent_reduce_values" => {
            super::list_adapter::emit_concurrent_reduce(chunks, current, 1, line);
        }
        "java.concurrent_reduce_entries" => {
            super::list_adapter::emit_concurrent_reduce(chunks, current, 2, line);
        }
        "java.concurrent_search_keys" => {
            super::list_adapter::emit_concurrent_search(chunks, current, 0, line);
        }
        "java.concurrent_search_values" => {
            super::list_adapter::emit_concurrent_search(chunks, current, 1, line);
        }
        "java.concurrent_search_entries" => {
            super::list_adapter::emit_concurrent_search(chunks, current, 2, line);
        }
        "java.semaphore_new" => {
            super::list_adapter::emit_semaphore_new(chunks, current, argc, line);
        }
        "java.semaphore_available" => {
            super::list_adapter::emit_semaphore_available(chunks, current, line);
        }
        "java.semaphore_acquire" => {
            super::list_adapter::emit_semaphore_acquire(chunks, current, argc, line);
        }
        "java.semaphore_release" => {
            super::list_adapter::emit_semaphore_release(chunks, current, argc, line);
        }
        "java.semaphore_try_acquire" => {
            super::list_adapter::emit_semaphore_try_acquire(chunks, current, argc, line);
        }
        "java.semaphore_drain" => {
            super::list_adapter::emit_semaphore_drain(chunks, current, line);
        }
        "java.semaphore_has_queued" => {
            super::list_adapter::emit_semaphore_has_queued(chunks, current, line);
        }
        "java.semaphore_queue_length" => {
            super::list_adapter::emit_semaphore_queue_length(chunks, current, line);
        }
        "java.semaphore_is_fair" => {
            super::list_adapter::emit_semaphore_is_fair(chunks, current, line);
        }
        "java.thread_start_with" => {
            super::list_adapter::emit_java_thread_start_with(chunks, current, line);
        }
        "java.thread_join" => {
            super::list_adapter::emit_java_thread_join(chunks, current, line);
        }
        "java.thread_sleep" => {
            super::list_adapter::emit_java_thread_sleep(chunks, current, line);
        }
        "java.get" => {
            if argc <= 1 {
                super::stream_adapter::emit_get_optional_value(chunks, current, line);
            } else {
                super::list_adapter::emit_get(chunks, current, line);
            }
        }
        "java.list_set" => {
            super::list_adapter::emit_set(chunks, current, argc, line);
        }
        "java.list_remove" => {
            super::list_adapter::emit_remove_at(chunks, current, line);
        }
        "java.list_remove_value" => {
            super::list_adapter::emit_remove_value_checked(chunks, current, line);
        }
        "java.list_clear" => {
            super::list_adapter::emit_clear(chunks, current, line);
        }
        "java.list_contains" => {
            collections::emit_contains(chunks, current, line);
        }
        "java.list_contains_all" => {
            super::list_adapter::emit_contains_all(chunks, current, line);
        }
        "java.list_equals" => {
            super::list_adapter::emit_list_equals(chunks, current, line);
        }
        "java.sub_list" => {
            super::list_adapter::emit_sub_list(chunks, current, line);
        }
        "java.list_sort" => {
            super::list_adapter::emit_sort(chunks, current, argc, line);
        }
        "java.add_all" => {
            super::list_adapter::emit_add_all(chunks, current, argc, line);
        }
        "java.remove_all" => {
            super::list_adapter::emit_remove_all(chunks, current, line);
        }
        "java.retain_all" => {
            super::list_adapter::emit_retain_all(chunks, current, line);
        }
        "java.list_remove_if" => {
            super::list_adapter::emit_remove_if(chunks, current, line);
        }
        "java.list_replace_all" => {
            super::list_adapter::emit_replace_all(chunks, current, line);
        }
        "java.list_for_each" => {
            // `list.forEach(action)` → shared HOF emitter (invokes `action`
            // per element; Java's `Consumer` ignores the extra index arg).
            let fn_slot = chunks[current].alloc_scratch(1);
            let arr_slot = chunks[current].alloc_scratch(1);
            let idx_slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
            vybe_compiler::compiler::loops::emit_foreach(chunks, current, fn_slot, arr_slot, idx_slot, line);
        }
        "java.queue_poll" => {
            super::list_adapter::emit_poll(chunks, current, false, line);
        }
        "java.priority_peek" => {
            super::list_adapter::emit_priority_peek(chunks, current, line);
        }
        "java.stack_push" => {
            super::list_adapter::emit_stack_push(chunks, current, line);
        }
        "java.stack_search" => {
            super::list_adapter::emit_stack_search(chunks, current, line);
        }
        "java.vector_capacity" => {
            super::list_adapter::emit_vector_capacity(chunks, current, line);
        }
        "java.vector_ensure_capacity" => {
            super::list_adapter::emit_vector_ensure_capacity(chunks, current, line);
        }
        "java.vector_trim_to_size" => {
            super::list_adapter::emit_vector_trim_to_size(chunks, current, line);
        }
        "java.vector_set_size" => {
            super::list_adapter::emit_vector_set_size(chunks, current, line);
        }
        "java.enumeration_from_array" => {
            super::list_adapter::emit_enumeration_from_array(chunks, current, line);
        }
        "java.enumeration_has_more" => {
            super::list_adapter::emit_enumeration_has_more(chunks, current, line);
        }
        "java.enumeration_next" => {
            super::list_adapter::emit_enumeration_next(chunks, current, line);
        }
        "java.hashtable_put" => {
            super::list_adapter::emit_hashtable_put(chunks, current, line);
        }
        "java.hashtable_keys" => {
            super::list_adapter::emit_hashtable_keys(chunks, current, line);
        }
        "java.hashtable_elements" => {
            super::list_adapter::emit_hashtable_elements(chunks, current, line);
        }
        "java.sorted_first" => {
            super::list_adapter::emit_sorted_end(chunks, current, false, line);
        }
        "java.sorted_last" => {
            super::list_adapter::emit_sorted_end(chunks, current, true, line);
        }
        "java.sorted_ceiling" => {
            super::list_adapter::emit_sorted_bound(chunks, current, 0, line);
        }
        "java.sorted_floor" => {
            super::list_adapter::emit_sorted_bound(chunks, current, 1, line);
        }
        "java.sorted_higher" => {
            super::list_adapter::emit_sorted_bound(chunks, current, 2, line);
        }
        "java.sorted_lower" => {
            super::list_adapter::emit_sorted_bound(chunks, current, 3, line);
        }
        "java.sorted_descending_set" => {
            super::list_adapter::emit_sorted_descending_set(chunks, current, line);
        }
        "java.sorted_sub_set" => {
            super::list_adapter::emit_sorted_set_range_view(chunks, current, 0, line);
        }
        "java.sorted_head_set" => {
            super::list_adapter::emit_sorted_set_range_view(chunks, current, 1, line);
        }
        "java.sorted_tail_set" => {
            super::list_adapter::emit_sorted_set_range_view(chunks, current, 2, line);
        }
        "java.sorted_first_key" => {
            super::list_adapter::emit_sorted_map_key(chunks, current, false, line);
        }
        "java.sorted_last_key" => {
            super::list_adapter::emit_sorted_map_key(chunks, current, true, line);
        }
        "java.add_first" => {
            host::emit(&mut chunks[current], "ecma:array", "unshift", 2, line);
        }
        "java.remove_first" => {
            collections::emit_shift(chunks, current, line);
        }
        "java.peek_first" => {
            super::list_adapter::emit_peek(chunks, current, false, line);
        }
        "java.poll_first" => {
            super::list_adapter::emit_poll(chunks, current, false, line);
        }
        "java.peek_last" => {
            super::list_adapter::emit_peek(chunks, current, true, line);
        }
        "java.poll_last" => {
            super::list_adapter::emit_poll(chunks, current, true, line);
        }

        // ── Map helpers ────────────────────────────────────────────────────
        "java.map_put" => {
            super::list_adapter::emit_map_put(chunks, current, line);
        }
        "java.map_put_all" => {
            super::list_adapter::emit_map_put_all(chunks, current, line);
        }
        "java.map_get_or_default" => {
            super::list_adapter::emit_map_get_or_default(chunks, current, line);
        }
        "java.map_contains_key" => {
            super::list_adapter::emit_map_contains_key(chunks, current, line);
        }
        "java.map_contains_value" => {
            super::list_adapter::emit_map_contains_value(chunks, current, line);
        }
        "java.map_key_set" => {
            super::list_adapter::emit_map_key_set(chunks, current, line);
        }
        "java.map_values" => {
            super::list_adapter::emit_map_values(chunks, current, line);
        }
        "java.entry_set" => {
            super::list_adapter::emit_map_entry_set(chunks, current, line);
        }
        "java.put_if_absent" => {
            super::list_adapter::emit_map_put_if_absent(chunks, current, line);
        }
        "java.compute_if_absent" => {
            super::list_adapter::emit_map_compute_if_absent(chunks, current, line);
        }
        "java.compute_if_present" => {
            super::list_adapter::emit_map_compute_if_present(chunks, current, line);
        }
        "java.map_compute" => {
            super::list_adapter::emit_map_compute(chunks, current, line);
        }
        "java.map_merge" => {
            super::list_adapter::emit_map_merge(chunks, current, line);
        }
        "java.map_remove" => {
            super::list_adapter::emit_map_remove(chunks, current, argc, line);
        }
        "java.map_replace" => {
            super::list_adapter::emit_map_replace(chunks, current, argc, line);
        }
        "java.map_replace_all" => {
            super::list_adapter::emit_map_replace_all(chunks, current, line);
        }
        "java.map_for_each" => {
            super::list_adapter::emit_map_for_each(chunks, current, line);
        }
        "java.map_clear" => {
            host::emit(&mut chunks[current], "ecma:map", "clear", 1, line);
        }
        "java.map_clone" => {
            super::list_adapter::emit_map_clone(chunks, current, line);
        }
        "java.map_size" => {
            host::emit(&mut chunks[current], "ecma:map", "size", 1, line);
        }
        "java.map_is_empty" => {
            host::emit(&mut chunks[current], "ecma:map", "size", 1, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op(Op::I32_EQ, line);
            vybe_compiler::compiler::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "java.map_equals" => {
            super::list_adapter::emit_map_equals(chunks, current, line);
        }
        "java.map_key_set_remove" => {
            super::list_adapter::emit_map_key_set_remove(chunks, current, line);
        }
        "java.sorted_map_key_set" => {
            super::list_adapter::emit_sorted_map_key_set(chunks, current, line);
        }
        "java.sorted_map_first_key" => {
            super::list_adapter::emit_sorted_map_key(chunks, current, false, line);
        }
        "java.sorted_map_last_key" => {
            super::list_adapter::emit_sorted_map_key(chunks, current, true, line);
        }
        "java.sorted_map_first_entry" => {
            super::list_adapter::emit_sorted_map_end_entry(chunks, current, false, line);
        }
        "java.sorted_map_last_entry" => {
            super::list_adapter::emit_sorted_map_end_entry(chunks, current, true, line);
        }
        "java.sorted_map_ceiling_entry" => {
            super::list_adapter::emit_sorted_map_bound_entry(chunks, current, 0, line);
        }
        "java.sorted_map_floor_entry" => {
            super::list_adapter::emit_sorted_map_bound_entry(chunks, current, 1, line);
        }
        "java.sorted_map_higher_entry" => {
            super::list_adapter::emit_sorted_map_bound_entry(chunks, current, 2, line);
        }
        "java.sorted_map_lower_entry" => {
            super::list_adapter::emit_sorted_map_bound_entry(chunks, current, 3, line);
        }
        "java.sorted_map_ceiling_key" => {
            super::list_adapter::emit_sorted_map_bound_key(chunks, current, 0, line);
        }
        "java.sorted_map_floor_key" => {
            super::list_adapter::emit_sorted_map_bound_key(chunks, current, 1, line);
        }
        "java.sorted_map_higher_key" => {
            super::list_adapter::emit_sorted_map_bound_key(chunks, current, 2, line);
        }
        "java.sorted_map_lower_key" => {
            super::list_adapter::emit_sorted_map_bound_key(chunks, current, 3, line);
        }
        "java.sorted_map_poll_first_entry" => {
            super::list_adapter::emit_sorted_map_poll_entry(chunks, current, false, line);
        }
        "java.sorted_map_poll_last_entry" => {
            super::list_adapter::emit_sorted_map_poll_entry(chunks, current, true, line);
        }
        "java.sorted_map_descending_key_set" => {
            super::list_adapter::emit_sorted_map_descending_key_set(chunks, current, line);
        }
        "java.sorted_map_descending_map" => {
            super::list_adapter::emit_sorted_map_descending_map(chunks, current, line);
        }
        "java.map_sub_map" => {
            super::list_adapter::emit_map_range_view(chunks, current, 0, line);
        }
        "java.map_head_map" => {
            super::list_adapter::emit_map_range_view(chunks, current, 1, line);
        }
        "java.map_tail_map" => {
            super::list_adapter::emit_map_range_view(chunks, current, 2, line);
        }
        "java.entry_key" => {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            collections::emit_get(chunks, current, line);
        }
        "java.entry_value" => {
            core_wasm::i32_const(&mut chunks[current], line, 1);
            collections::emit_get(chunks, current, line);
        }
        "java.entry_set_value" => {
            super::list_adapter::emit_entry_set_value(chunks, current, line);
        }
        "java.list_iterator" => {
            super::list_adapter::emit_list_iterator(chunks, current, argc, line);
        }
        "java.iterator_next" => {
            super::list_adapter::emit_iterator_next(chunks, current, line);
        }
        "java.iterator_has_next" => {
            super::list_adapter::emit_iterator_has_next(chunks, current, line);
        }
        "java.iterator_previous" => {
            super::list_adapter::emit_iterator_previous(chunks, current, line);
        }
        "java.iterator_has_previous" => {
            super::list_adapter::emit_iterator_has_previous(chunks, current, line);
        }
        "java.iterator_next_index" => {
            super::list_adapter::emit_iterator_next_index(chunks, current, line);
        }
        "java.iterator_previous_index" => {
            super::list_adapter::emit_iterator_previous_index(chunks, current, line);
        }

        // ── StringBuilder helpers ──────────────────────────────────────────
        "java.stringbuilder_new" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.string_builder_new",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_append" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.sb_append",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_insert" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.sb_insert",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_delete" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.sb_remove",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_delete_char_at" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.sb_remove",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_reverse" => {
            collections::emit_reverse(chunks, current, line);
        }

        // ── Collections utilities ─────────────────────────────────────────
        "java.collections_sort" => {
            if argc == 2 {
                collections::emit_sort_with_comparator(chunks, current, line);
            } else {
                collections::emit_sort(chunks, current, line);
            }
        }
        "java.collections_reverse" => {
            collections::emit_reverse(chunks, current, line);
        }
        "java.collections_shuffle" => {
            if argc == 2 {
                chunks[current].emit_op(Op::DROP, line);
            }
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op(Op::NULL, line);
        }
        "java.collections_fill" => {
            super::arrays_adapter::emit_fill(chunks, current, 2, line);
        }
        "java.collections_copy" => {
            super::list_adapter::emit_collection_copy(chunks, current, line);
        }
        "java.collections_min" => {
            super::list_adapter::emit_collection_extreme(chunks, current, argc, true, line);
        }
        "java.collections_max" => {
            super::list_adapter::emit_collection_extreme(chunks, current, argc, false, line);
        }
        "java.collections_frequency" => {
            super::list_adapter::emit_collection_frequency(chunks, current, line);
        }
        "java.collections_disjoint" => {
            super::list_adapter::emit_collection_disjoint(chunks, current, line);
        }

        // ── String formatting ─────────────────────────────────────────────
        "java.string_format" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.string_format",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.string_join" => {
            super::string_adapter::emit_join(chunks, current, argc, line);
        }

        // ── Optional ─────────────────────────────────────────────────────
        "java.optional_empty" => {
            super::optional_adapter::emit_empty(chunks, current, line);
        }
        "java.optional_of" => {
            super::optional_adapter::emit_of(chunks, current, line);
        }
        "java.optional_of_long" => {
            super::optional_adapter::emit_of_long(chunks, current, line);
        }
        "java.optional_of_nullable" => {
            super::optional_adapter::emit_of_nullable(chunks, current, line);
        }
        "java.optional_or_else" | "java.optional_or_else_get" => {
            super::optional_adapter::emit_or_else(
                chunks,
                current,
                name == "java.optional_or_else_get",
                line,
            );
        }
        "java.optional_is_present" => {
            super::optional_adapter::emit_is_present(chunks, current, line);
        }
        "java.optional_if_present" => {
            super::optional_adapter::emit_if_present(chunks, current, line);
        }
        "java.optional_filter" => {
            super::optional_adapter::emit_filter(chunks, current, line);
        }
        "java.optional_map" => {
            super::optional_adapter::emit_map(chunks, current, line);
        }
        "java.optional_flat_map" => {
            super::optional_adapter::emit_flat_map(chunks, current, line);
        }
        "java.optional_if_present_or_else" => {
            super::optional_adapter::emit_if_present_or_else(chunks, current, line);
        }
        "java.optional_is_empty" => {
            super::optional_adapter::emit_is_empty(chunks, current, line);
        }
        "java.optional_or" | "java.optional_or_get" => {
            super::optional_adapter::emit_or(chunks, current, name == "java.optional_or_get", line);
        }
        "java.optional_stream" => {
            super::optional_adapter::emit_stream(chunks, current, line);
        }
        "java.optional_equals" => {
            super::optional_adapter::emit_equals(chunks, current, line);
        }
        "java.optional_to_string" => {
            super::optional_adapter::emit_to_string(chunks, current, line);
        }
        "java.optional_or_else_throw" => {
            super::optional_adapter::emit_or_else_throw(chunks, current, argc > 1, line);
        }

        // ── Object utilities ──────────────────────────────────────────────
        "java.equals" => {
            vybe_compiler::compiler::object::emit_equals(&mut chunks[current], line);
        }
        "java.objects_equals" => {
            vybe_compiler::compiler::object::emit_equals(&mut chunks[current], line);
        }
        "java.hash_code" => {
            vybe_compiler::compiler::object::emit_hash_code(&mut chunks[current], line);
        }
        "java.require_non_null" => {
            if argc > 1 {
                chunks[current].emit_op(Op::DROP, line);
            }
        }
        "java.is_null" => {
            vybe_compiler::compiler::object::emit_is_null(&mut chunks[current], line);
        }
        "java.non_null" => {
            vybe_compiler::compiler::object::emit_non_null(&mut chunks[current], line);
        }

        _ => return false,
    }
    true
}

enum JavaArrayDefault {
    IntZero,
    BoolFalse,
}

fn emit_new_array_with_default(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    default: JavaArrayDefault,
) {
    let len_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    collections::emit_new_with_length(chunks, current, line);

    let arr_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    match default {
        JavaArrayDefault::IntZero => chunks[current].emit_i32_const(0, line),
        JavaArrayDefault::BoolFalse => chunks[current].emit_bool_const(false, line),
    }
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    collections::emit_fill(chunks, current, line);
}
