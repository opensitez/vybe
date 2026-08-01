//! `jvm.*` emit dispatch — the platform's own adapter routing.
//!
//! Mirrors `platforms/dotnet`: an op named `common:jvm.java.<name>` reaches this
//! function through `primitives::dispatch` → `platform_emit_dispatch_for`, so
//! the JDK's behaviour is emitted from the PLATFORM and every JVM language
//! gets it by resolving the tree — no per-language emitter arms, no prelude.

use vybe_compiler::primitives::url::UrlField;
use vybe_compiler::primitives::{
    instructions::{core_wasm, host},
    ops, strings,
};
use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    use crate::emitter::arrays_adapter as arrays;
    use crate::emitter::bitset_adapter as bitset;
    use crate::emitter::map_adapter as map;
    use crate::emitter::math_adapter as math;
    use crate::emitter::stringbuilder_adapter as sb;
    use crate::emitter::stringtokenizer_adapter as st;
    use crate::emitter::system_adapter as system;
    use crate::emitter::url_adapter as url;
    use crate::emitter::uuid_adapter as uuid;
    match name {
        // ── construction ──
        "jvm.java.net.url_new" => url::emit_url_new(chunks, current, argc, line),
        "jvm.java.net.uri_new" => url::emit_uri_new(chunks, current, argc, line),

        // ── components: the SHARED reader in primitives/url.rs ──
        "jvm.java.net.url_protocol" => {
            url::emit_component_getter(chunks, current, UrlField::Scheme, line)
        }
        "jvm.java.net.url_host" => {
            url::emit_component_getter(chunks, current, UrlField::Host, line)
        }
        "jvm.java.net.url_path" => {
            url::emit_component_getter(chunks, current, UrlField::Path, line)
        }
        "jvm.java.net.url_authority" => {
            url::emit_component_getter(chunks, current, UrlField::Netloc, line)
        }
        "jvm.java.net.url_query" => {
            url::emit_nullable_getter(chunks, current, UrlField::Query, line)
        }
        "jvm.java.net.url_ref" => {
            url::emit_nullable_getter(chunks, current, UrlField::Fragment, line)
        }
        "jvm.java.net.url_port" => url::emit_port(chunks, current, line),
        "jvm.java.net.url_default_port" => url::emit_default_port(chunks, current, line),
        "jvm.java.net.url_file" => url::emit_file(chunks, current, line),
        "jvm.java.net.url_user_info" => url::emit_user_info(chunks, current, line),

        // ── identity and text ──
        "jvm.java.net.url_to_string" => url::emit_to_string(chunks, current, line),
        "jvm.java.net.url_to_uri" => url::emit_url_to_uri(chunks, current, line),
        "jvm.java.net.url_equals" => url::emit_equals(chunks, current, line),
        "jvm.java.net.url_hash" => url::emit_hash(chunks, current, line),
        "jvm.java.net.url_same_file" => url::emit_same_file(chunks, current, line),
        "jvm.java.net.url_encode" => url::emit_url_encode(chunks, current, line),
        "jvm.java.net.url_decode" => url::emit_url_decode(chunks, current, line),

        // ── java.net.URI's relational surface ──
        "jvm.java.net.uri_ssp" => url::emit_ssp(chunks, current, line),
        "jvm.java.net.uri_is_absolute" => url::emit_is_absolute(chunks, current, line),
        "jvm.java.net.uri_is_opaque" => url::emit_is_opaque(chunks, current, line),
        "jvm.java.net.uri_to_url" => url::emit_uri_to_url(chunks, current, line),
        "jvm.java.net.uri_normalize" => url::emit_normalize(chunks, current, line),
        "jvm.java.net.uri_resolve" => url::emit_resolve(chunks, current, line),
        "jvm.java.net.uri_relativize" => url::emit_relativize(chunks, current, line),
        "jvm.java.net.uri_compare_to" => url::emit_compare_to(chunks, current, line),
        "jvm.java.lang.system_get_property" => {
            system::emit_get_property(chunks, current, argc, line);
        }
        "jvm.java.uuid_from_string" => uuid::emit_from_string(chunks, current, line),
        "jvm.java.uuid_name_from_bytes" => uuid::emit_name_from_bytes(chunks, current, line),
        "jvm.java.uuid_version" => uuid::emit_version(chunks, current, line),
        "jvm.java.uuid_variant" => uuid::emit_variant(chunks, current, line),
        "jvm.java.uuid_most_bits" => uuid::emit_most_bits(chunks, current, line),
        "jvm.java.uuid_least_bits" => uuid::emit_least_bits(chunks, current, line),
        "jvm.java.uuid_compare_to" => uuid::emit_compare_to(chunks, current, line),
        "jvm.java.uuid_hash_code" => uuid::emit_hash_code(chunks, current, line),
        "jvm.java.uuid_new" => uuid::emit_new(chunks, current, argc, line),
        "jvm.java.is_infinite" => {
            host::emit(&mut chunks[current], "ecma:number", "isFinite", 1, line);
            ops::emit_dyn_not(&mut chunks[current], line);
        }
        "jvm.java.signum" => host::emit(&mut chunks[current], "ecma:math", "sign", 1, line),
        "jvm.java.math_scalb" => math::emit_scalb(chunks, current, line),
        "jvm.java.math_ulp" => math::emit_ulp(chunks, current, line),
        "jvm.java.math_get_exponent" => math::emit_get_exponent(chunks, current, line),
        "jvm.java.math_copy_sign" => math::emit_copy_sign(chunks, current, line),
        "jvm.java.math_next_after" => math::emit_next_after(chunks, current, line),
        "jvm.java.math_next_up" => math::emit_next_up(chunks, current, line),
        "jvm.java.math_next_down" => math::emit_next_down(chunks, current, line),
        "jvm.java.math_fma" => math::emit_fma(chunks, current, line),
        "jvm.java.math_expm1" => math::emit_expm1(chunks, current, line),
        "jvm.java.math_log1p" => math::emit_log1p(chunks, current, line),
        "jvm.java.math_to_degrees" => math::emit_to_degrees(chunks, current, line),
        "jvm.java.math_to_radians" => math::emit_to_radians(chunks, current, line),
        "jvm.java.math_ieee_remainder" => math::emit_ieee_remainder(chunks, current, line),
        "jvm.java.math_add_exact" => math::emit_add_exact(chunks, current, line),
        "jvm.java.math_subtract_exact" => math::emit_subtract_exact(chunks, current, line),
        "jvm.java.math_multiply_exact" => math::emit_multiply_exact(chunks, current, line),
        "jvm.java.math_increment_exact" => math::emit_increment_exact(chunks, current, line),
        "jvm.java.math_decrement_exact" => math::emit_decrement_exact(chunks, current, line),
        "jvm.java.math_negate_exact" => math::emit_negate_exact(chunks, current, line),
        "jvm.java.floor_div" => math::emit_floor_div(chunks, current, line),
        "jvm.java.floor_mod" => math::emit_floor_mod(chunks, current, line),
        "jvm.java.identity" => {}
        "jvm.java.char_is_digit" => {
            host::emit(&mut chunks[current], "ecma:char", "isDigit", 1, line);
        }
        "jvm.java.char_is_letter" => {
            host::emit(&mut chunks[current], "ecma:char", "isLetter", 1, line);
        }
        "jvm.java.char_is_alnum" => {
            host::emit(&mut chunks[current], "ecma:char", "isAlnum", 1, line);
        }
        "jvm.java.char_is_upper" => {
            host::emit(&mut chunks[current], "ecma:char", "isUpper", 1, line);
        }
        "jvm.java.char_is_lower" => {
            host::emit(&mut chunks[current], "ecma:char", "isLower", 1, line);
        }
        "jvm.java.char_is_space" => {
            host::emit(&mut chunks[current], "ecma:char", "isSpace", 1, line);
        }
        "jvm.java.char_to_upper" => strings::emit_to_upper(&mut chunks[current], line),
        "jvm.java.char_to_lower" => strings::emit_to_lower(&mut chunks[current], line),
        "jvm.java.char_numeric" => {
            host::emit(&mut chunks[current], "ecma:number", "parseInt", 1, line);
        }
        "jvm.java.to_binary_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toBinary", 1, line);
        }
        "jvm.java.to_hex_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toHex", 1, line);
        }
        "jvm.java.to_octal_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toOctal", 1, line);
        }
        "jvm.java.parse_int" => {
            host::emit(&mut chunks[current], "ecma:number", "parseInt", argc, line);
        }
        "jvm.java.int_bit_count" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_POPCNT, line);
        }
        "jvm.java.int_leading_zeros" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_CLZ, line);
        }
        "jvm.java.int_trailing_zeros" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_CTZ, line);
        }
        "jvm.java.int_rotate_left" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_ROTL, line);
        }
        "jvm.java.int_rotate_right" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_ROTR, line);
        }
        "jvm.java.int_lowest_one_bit" => {
            let s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, s, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_SUB, line);
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_AND, line);
        }
        "jvm.java.int_highest_one_bit" => {
            let s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, s, line);
            for shift in [1, 2, 4, 8, 16] {
                chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
                chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
                core_wasm::i32_const(&mut chunks[current], line, shift);
                chunks[current].emit_op(vybe_runtime::opcode::Op::I32_SHR_U, line);
                chunks[current].emit_op(vybe_runtime::opcode::Op::I32_OR, line);
                chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, s, line);
            }
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_SHR_U, line);
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_SUB, line);
        }
        "jvm.java.compare" => {
            let b_slot = chunks[current].alloc_scratch(1);
            let a_slot = chunks[current].alloc_scratch(1);
            let result_slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, b_slot, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, a_slot, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, a_slot, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, b_slot, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(-1, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, a_slot, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, b_slot, line);
            ops::emit_dyn_gt(&mut chunks[current], line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, result_slot, line);
        }
        "jvm.java.arrays_sort" => arrays::emit_sort(chunks, current, argc, line),
        "jvm.java.arrays_fill" => arrays::emit_fill(chunks, current, argc, line),
        "jvm.java.arrays_copy_of" => arrays::emit_copy_of(chunks, current, line),
        "jvm.java.arrays_copy_of_range" => arrays::emit_copy_of_range(chunks, current, line),
        "jvm.java.arrays_to_string" => arrays::emit_to_string(chunks, current, line),
        "jvm.java.arrays_deep_to_string" => arrays::emit_deep_to_string(chunks, current, line),
        "jvm.java.arrays_equals" => arrays::emit_equals(chunks, current, line),
        "jvm.java.arrays_deep_equals" => arrays::emit_deep_equals(chunks, current, line),
        "jvm.java.arrays_compare" => arrays::emit_compare(chunks, current, line),
        "jvm.java.arrays_compare_unsigned" => arrays::emit_compare_unsigned(chunks, current, line),
        "jvm.java.arrays_mismatch" => arrays::emit_mismatch(chunks, current, line),
        "jvm.java.arrays_set_all" => arrays::emit_set_all(chunks, current, line),
        "jvm.java.arrays_parallel_prefix" => {
            arrays::emit_parallel_prefix(chunks, current, argc, line)
        }
        "jvm.java.arrays_binary_search" => arrays::emit_binary_search(chunks, current, argc, line),
        "jvm.java.arrays_as_list" => arrays::emit_arrays_as_list(chunks, current, argc, line),
        "jvm.java.arrays_hash_code" => arrays::emit_hash_code(chunks, current, line),
        "jvm.java.arrays_deep_hash_code" => arrays::emit_deep_hash_code(chunks, current, line),
        "jvm.java.bitset_new" => bitset::emit_new(chunks, current, argc, line),
        "jvm.java.bitset_value_of" => bitset::emit_value_of(chunks, current, line),
        "jvm.java.bitset_set" => bitset::emit_set(chunks, current, argc, line),
        "jvm.java.bitset_get" => bitset::emit_get(chunks, current, argc, line),
        "jvm.java.bitset_clear" => bitset::emit_clear(chunks, current, argc, line),
        "jvm.java.bitset_flip" => bitset::emit_flip(chunks, current, argc, line),
        "jvm.java.bitset_cardinality" => bitset::emit_cardinality(chunks, current, line),
        "jvm.java.bitset_length" => bitset::emit_length(chunks, current, line),
        "jvm.java.bitset_size" => bitset::emit_size(chunks, current, line),
        "jvm.java.bitset_is_empty" => bitset::emit_is_empty(chunks, current, line),
        "jvm.java.bitset_next_set_bit" => bitset::emit_next_set_bit(chunks, current, line),
        "jvm.java.bitset_next_clear_bit" => bitset::emit_next_clear_bit(chunks, current, line),
        "jvm.java.bitset_previous_set_bit" => bitset::emit_previous_set_bit(chunks, current, line),
        "jvm.java.bitset_previous_clear_bit" => {
            bitset::emit_previous_clear_bit(chunks, current, line)
        }
        "jvm.java.bitset_and" => bitset::emit_and(chunks, current, line),
        "jvm.java.bitset_or" => bitset::emit_or(chunks, current, line),
        "jvm.java.bitset_xor" => bitset::emit_xor(chunks, current, line),
        "jvm.java.bitset_and_not" => bitset::emit_and_not(chunks, current, line),
        "jvm.java.bitset_intersects" => bitset::emit_intersects(chunks, current, line),
        "jvm.java.bitset_equals" => bitset::emit_equals(chunks, current, line),
        "jvm.java.bitset_clone" => bitset::emit_clone(chunks, current, line),
        "jvm.java.bitset_stream" => bitset::emit_stream(chunks, current, line),
        "jvm.java.bitset_to_array" => bitset::emit_to_array(chunks, current, line),
        "jvm.java.bitset_to_string" => bitset::emit_to_string(chunks, current, line),
        "jvm.java.bitset_hash_code" => bitset::emit_hash_code(chunks, current, line),
        "jvm.java.hash_map_new" => map::emit_hash_map_new(chunks, current, argc, line),
        "jvm.java.concurrent_hash_map_new" => {
            map::emit_concurrent_hash_map_new(chunks, current, argc, line)
        }
        "jvm.java.identity_hash_map_new" => {
            map::emit_identity_hash_map_new(chunks, current, argc, line)
        }
        "jvm.java.linked_hash_map_new" => {
            map::emit_linked_hash_map_new(chunks, current, argc, line)
        }
        "jvm.java.stringbuilder_new" => sb::emit_new(chunks, current, argc, line),
        "jvm.java.sb_append" => sb::emit_append(chunks, current, argc, line),
        "jvm.java.sb_append_line" => sb::emit_append_line(chunks, current, argc, line),
        "jvm.java.sb_to_string" => sb::emit_to_string(chunks, current, argc, line),
        "jvm.java.stringtokenizer_new" => st::emit_new(chunks, current, argc, line),
        "jvm.java.st_has_more" => st::emit_has_more(chunks, current, argc, line),
        "jvm.java.st_next" => st::emit_next(chunks, current, argc, line),
        "jvm.java.st_count" => st::emit_count(chunks, current, argc, line),

        _ => return false,
    }
    true
}
