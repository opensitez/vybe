//! PHP `common:php.<name>` dispatch.
//!
//! Routes PHP profile emit keys to the PHP runtime adapters in this module.
//! The shared `emitter::dispatch::emit_common` delegates every `php.*` key
//! here (with a fall-through to the central table for keys not yet migrated),
//! so PHP-specific routing lives in the PHP module instead of the common
//! dispatcher. Returns `true` if `name` was recognized and emitted.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// PHP `isset($a, $b, ...)` — true iff every arg is non-null. Variadic, so it
/// can't be a fixed-arity stdlib chunk.
fn emit_isset_all(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc == 0 {
        chunk.emit_bool_const(true, line);
        return;
    }
    let base = chunk.local_count;
    chunk.local_count = base + argc as u16;
    for i in (0..argc).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    for i in 1..argc {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        vybe_emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_op(Op::I32_AND, line);
    }
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        name if super::runtime_adapter::emit_helper(name, chunks, current, argc, line) => {}
        // ── PHP array helpers ──────────────────────────────────────
        // Index-based loops + ECMA array/object ops. PHP `array` ≡
        // `Map` (assoc) or `Array` (sequential).
        "php.echo" => super::output_adapter::emit_php_echo(chunks, current, argc, line),
        "php.print_expr" => super::output_adapter::emit_php_print_expr(chunks, current, line),
        "php.compare_gt" => {
            let chunk = &mut chunks[current];
            super::relational_adapter::emit_relational_compare(
                chunk,
                vybe_emitter::ops::emit_dyn_gt,
                line,
            );
        }
        "php.compare_lt" => {
            let chunk = &mut chunks[current];
            super::relational_adapter::emit_relational_compare(
                chunk,
                vybe_emitter::ops::emit_dyn_lt,
                line,
            );
        }
        "php.compare_gte" => {
            let chunk = &mut chunks[current];
            super::relational_adapter::emit_relational_compare(
                chunk,
                vybe_emitter::ops::emit_dyn_ge,
                line,
            );
        }
        "php.compare_lte" => {
            let chunk = &mut chunks[current];
            super::relational_adapter::emit_relational_compare(
                chunk,
                vybe_emitter::ops::emit_dyn_le,
                line,
            );
        }
        "php.intval" => {
            let num_idx = chunks[current].add_import("ecma:number", "Number");
            chunks[current].emit_call(num_idx, 1, line);
            chunks[current].emit_op(vybe_bytecode::opcode::Op::F64_TRUNC, line);
        }
        // Guarded delegators: PHP throws TypeError when the first argument
        // isn't an array; the underlying op is the shared `common:*` emit.
        "php.array_push_g" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "array_push(): Argument #1 ($array) must be of type array",
                line,
            );
            vybe_emitter::collections::emit_push(chunks, current, line);
        }
        "php.array_pop_g" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "array_pop(): Argument #1 ($array) must be of type array",
                line,
            );
            vybe_emitter::collections::emit_pop(chunks, current, line);
        }
        "php.sort_g" => {
            // `sort($x)` on a non-array throws (base `\Error` in PHP because
            // the parameter is by-reference; `TypeError` is an `\Error`
            // subclass, so `catch (\Error)` matches either way).
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "sort(): Argument #1 ($array) must be of type array",
                line,
            );
            super::runtime_adapter::emit_helper("php.sort_in_place", chunks, current, argc, line);
        }
        "php.array_pad" => super::array_adapter::emit_array_pad(chunks, current, argc, line),
        "php.array_map" => super::array_adapter::emit_array_map(chunks, current, argc, line),
        "php.array_filter" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "array_filter(): Argument #1 ($array) must be of type array",
                line,
            );
            super::array_adapter::emit_array_filter(chunks, current, argc, line)
        }
        "php.array_walk_recursive" => {
            super::array_adapter::emit_array_walk_recursive(chunks, current, argc, line)
        }
        "php.array_fill" => super::array_adapter::emit_array_fill(chunks, current, argc, line),
        "php.array_fill_keys" => {
            super::array_adapter::emit_array_fill_keys(chunks, current, argc, line)
        }
        "php.array_search" => super::array_adapter::emit_array_search(chunks, current, argc, line),
        "php.array_key_exists" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "array_key_exists(): Argument #2 ($array) must be of type array",
                line,
            );
            super::array_adapter::emit_array_key_exists(chunks, current, argc, line)
        }
        "php.count" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::NotScalar,
                "TypeError",
                "count(): Argument #1 ($value) must be of type Countable|array",
                line,
            );
            super::array_adapter::emit_php_count(chunks, current, argc, line)
        }
        "php.json_encode" => {
            super::array_adapter::emit_php_json_encode(chunks, current, argc, line)
        }
        "php.json_decode" => {
            super::array_adapter::emit_php_json_decode(chunks, current, argc, line)
        }
        "php.json_last_error" => {
            super::array_adapter::emit_php_json_last_error(chunks, current, argc, line)
        }
        "php.json_last_error_msg" => {
            super::array_adapter::emit_php_json_last_error_msg(chunks, current, argc, line)
        }
        "php.json_validate" => {
            super::array_adapter::emit_php_json_validate(chunks, current, argc, line)
        }
        "php.simplexml_load_string" => {
            super::xml_adapter::emit_simplexml_load_string(chunks, current, argc, line)
        }
        "php.dom_save_xml" => super::xml_adapter::emit_dom_save_xml(chunks, current, argc, line),
        "php.array_keys" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "array_keys(): Argument #1 ($array) must be of type array",
                line,
            );
            super::array_adapter::emit_php_array_keys(chunks, current, argc, line)
        }
        "php.array_values" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "array_values(): Argument #1 ($array) must be of type array",
                line,
            );
            super::array_adapter::emit_php_array_values(chunks, current, argc, line)
        }
        "php.array_is_list" => {
            super::array_adapter::emit_php_array_is_list(chunks, current, argc, line)
        }
        "php.end" => super::array_adapter::emit_php_end(chunks, current, argc, line),
        "php.array_chunk" => super::array_adapter::emit_array_chunk(chunks, current, argc, line),
        "php.array_combine" => {
            super::array_adapter::emit_array_combine(chunks, current, argc, line)
        }
        "php.array_flip" => super::array_adapter::emit_array_flip(chunks, current, argc, line),
        "php.array_diff" => super::array_adapter::emit_array_diff(chunks, current, argc, line),
        "php.array_intersect" => {
            super::array_adapter::emit_array_intersect(chunks, current, argc, line)
        }
        "php.array_count_values" => {
            super::array_adapter::emit_array_count_values(chunks, current, argc, line)
        }
        "php.array_column" => super::array_adapter::emit_array_column(chunks, current, argc, line),
        "php.array_key_first" => {
            super::array_adapter::emit_array_key_first(chunks, current, argc, line)
        }
        "php.array_key_last" => {
            super::array_adapter::emit_array_key_last(chunks, current, argc, line)
        }
        "php.array_diff_key" => {
            super::array_adapter::emit_array_diff_key(chunks, current, argc, line)
        }
        "php.array_diff_assoc" => {
            super::array_adapter::emit_array_diff_assoc(chunks, current, argc, line)
        }
        "php.array_intersect_assoc" => {
            super::array_adapter::emit_array_intersect_assoc(chunks, current, argc, line)
        }
        "php.array_intersect_key" => {
            super::array_adapter::emit_array_intersect_key(chunks, current, argc, line)
        }
        "php.array_intersect_ukey" => {
            super::array_adapter::emit_array_intersect_ukey(chunks, current, argc, line)
        }
        "php.array_udiff_uassoc" => {
            super::array_adapter::emit_array_udiff_uassoc(chunks, current, argc, line)
        }
        "php.array_uintersect_uassoc" => {
            super::array_adapter::emit_array_uintersect_uassoc(chunks, current, argc, line)
        }
        "php.array_replace" => {
            super::array_adapter::emit_array_replace(chunks, current, argc, line)
        }
        "php.array_replace_recursive" => {
            super::array_adapter::emit_array_replace_recursive(chunks, current, argc, line)
        }
        "php.iterator_to_array" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::NotScalar,
                "TypeError",
                "iterator_to_array(): Argument #1 ($iterator) must be of type Traversable|array",
                line,
            );
            super::array_adapter::emit_iterator_to_array(chunks, current, argc, line)
        }
        "php.generator_key" => {
            super::array_adapter::emit_generator_key(chunks, current, argc, line)
        }
        "php.generator_get_return" => {
            super::array_adapter::emit_generator_get_return(chunks, current, argc, line)
        }
        "php.generator_rewind" => {
            super::array_adapter::emit_generator_rewind(chunks, current, argc, line)
        }
        "php.generator_next" => {
            super::array_adapter::emit_generator_next(chunks, current, argc, line)
        }
        "php.generator_current" => {
            super::array_adapter::emit_generator_current(chunks, current, argc, line)
        }
        "php.generator_valid" => {
            super::array_adapter::emit_generator_valid(chunks, current, argc, line)
        }
        "php.asort" => super::array_adapter::emit_php_asort(chunks, current, argc, line),
        "php.arsort" => super::array_adapter::emit_php_arsort(chunks, current, argc, line),
        "php.ksort" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "ksort(): Argument #1 ($array) must be of type array",
                line,
            );
            super::array_adapter::emit_php_ksort(chunks, current, argc, line)
        }
        "php.krsort" => super::array_adapter::emit_php_krsort(chunks, current, argc, line),
        "php.uasort" => super::array_adapter::emit_php_uasort(chunks, current, argc, line),
        "php.uksort" => super::array_adapter::emit_php_uksort(chunks, current, argc, line),
        // NOTE: `implode([1,2], '-')` should throw in PHP 8 (the array-first
        // arg order was removed), but the walker (walker.rs ~9445) still
        // unconditionally swaps 2-arg implode, which both hides the error and
        // would misfire on the valid `implode('-', [1,2])` order. A guard here
        // can't distinguish the two post-normalization — the fix belongs in the
        // walker. Left unguarded for now.
        "php.implode" => super::array_adapter::emit_php_implode(chunks, current, argc, line),
        "php.in_array" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "in_array(): Argument #2 ($haystack) must be of type array",
                line,
            );
            super::array_adapter::emit_php_in_array(chunks, current, argc, line)
        }
        "php.obj_to_array" => {
            super::array_adapter::emit_php_obj_to_array(chunks, current, argc, line)
        }
        "php.array_to_object" => {
            super::array_adapter::emit_php_array_to_object(chunks, current, argc, line)
        }
        "php.var_export" => {
            crate::emitter::string_adapter::emit_var_export(chunks, current, argc, line)
        }
        "php.print_r" => super::array_adapter::emit_php_print_r(chunks, current, argc, line),

        "php.datetime_new" => {
            crate::emitter::datetime_adapter::emit_datetime_new(chunks, current, argc, line)
        }
        "php.datetime_immutable_new" => {
            crate::emitter::datetime_adapter::emit_datetime_immutable_new(
                chunks, current, argc, line,
            )
        }
        "php.datetimezone_new" => {
            crate::emitter::datetime_adapter::emit_datetimezone_new(chunks, current, line)
        }
        "php.datetime_get_timezone" => {
            crate::emitter::datetime_adapter::emit_datetime_get_timezone(chunks, current, line)
        }
        "php.datetime_get_offset" => {
            crate::emitter::datetime_adapter::emit_datetime_get_offset(chunks, current, argc, line)
        }
        "php.datetime_set_timezone" => {
            crate::emitter::datetime_adapter::emit_datetime_set_timezone(chunks, current, line)
        }
        "php.datetime_set_date" => {
            crate::emitter::datetime_adapter::emit_datetime_set_date(chunks, current, line)
        }
        "php.datetime_set_time" => {
            crate::emitter::datetime_adapter::emit_datetime_set_time(chunks, current, line)
        }
        "php.datetime_create_from_format" => {
            crate::emitter::datetime_adapter::emit_datetime_create_from_format(
                chunks, current, line,
            )
        }
        "php.datetime_immutable_create_from_format" => {
            crate::emitter::datetime_adapter::emit_datetime_immutable_create_from_format(
                chunks, current, line,
            )
        }
        "php.datetime_format" => {
            crate::emitter::datetime_adapter::emit_datetime_format(chunks, current, line)
        }
        "php.datetime_get_timestamp" => {
            crate::emitter::datetime_adapter::emit_datetime_get_timestamp(chunks, current, line)
        }
        "php.datetime_modify" => {
            crate::emitter::datetime_adapter::emit_datetime_modify(chunks, current, line)
        }
        "php.datetime_immutable_modify" => {
            crate::emitter::datetime_adapter::emit_datetime_immutable_modify(chunks, current, line)
        }
        "php.datetime_diff" => {
            crate::emitter::datetime_adapter::emit_datetime_diff(chunks, current, argc, line)
        }
        "php.datetime_add" => {
            crate::emitter::datetime_adapter::emit_datetime_add(chunks, current, line)
        }
        "php.datetime_sub" => {
            crate::emitter::datetime_adapter::emit_datetime_sub(chunks, current, line)
        }
        "php.datetime_immutable_add" => {
            crate::emitter::datetime_adapter::emit_datetime_immutable_add(chunks, current, line)
        }
        "php.datetime_immutable_sub" => {
            crate::emitter::datetime_adapter::emit_datetime_immutable_sub(chunks, current, line)
        }
        "php.dateinterval_components" => {
            crate::emitter::datetime_adapter::emit_dateinterval_components(chunks, current, line)
        }
        "php.datetime_add_seconds" => {
            crate::emitter::datetime_adapter::emit_datetime_add_seconds(chunks, current, line)
        }
        "php.datetime_add_minutes" => {
            crate::emitter::datetime_adapter::emit_datetime_add_minutes(chunks, current, line)
        }
        "php.datetime_add_hours" => {
            crate::emitter::datetime_adapter::emit_datetime_add_hours(chunks, current, line)
        }
        "php.datetime_add_days" => {
            crate::emitter::datetime_adapter::emit_datetime_add_days(chunks, current, line)
        }
        "php.datetime_add_weeks" => {
            crate::emitter::datetime_adapter::emit_datetime_add_weeks(chunks, current, line)
        }
        "php.datetime_add_months" => {
            crate::emitter::datetime_adapter::emit_datetime_add_months(chunks, current, line)
        }
        "php.datetime_add_years" => {
            crate::emitter::datetime_adapter::emit_datetime_add_years(chunks, current, line)
        }

        // ── PHP top-level date functions — Rust opcode emitters ────
        // Compose `ecma:date.now/parse/UTC/get*` into PHP-shaped
        // surfaces: `date()`, `strftime()`, `strtotime()`, `mktime()`.
        // Pure bytecode; no PHP-specific host fns.
        "php.date" => crate::emitter::datetime_adapter::emit_php_date(chunks, current, argc, line),
        "php.strftime" => {
            crate::emitter::datetime_adapter::emit_php_strftime(chunks, current, argc, line)
        }
        "php.strtotime" => {
            crate::emitter::datetime_adapter::emit_php_strtotime(chunks, current, argc, line)
        }
        "php.strtotime_rel_calendar" => {
            crate::emitter::datetime_adapter::emit_php_strtotime_rel_calendar(
                chunks, current, argc, line,
            )
        }
        "php.mktime" => {
            crate::emitter::datetime_adapter::emit_php_mktime(chunks, current, argc, line)
        }
        "php.checkdate" => {
            crate::emitter::datetime_adapter::emit_php_checkdate(chunks, current, argc, line)
        }
        "php.getdate" => {
            crate::emitter::datetime_adapter::emit_php_getdate(chunks, current, argc, line)
        }

        // ── PHP `$x++` / `$x--` arithmetic ─────────────────────────
        // Composes `ecma:number.parseFloat` for string-numeric coerce.
        "php.inc" => crate::emitter::numeric_adapter::emit_php_inc(chunks, current, argc, line),
        "php.dec" => crate::emitter::numeric_adapter::emit_php_dec(chunks, current, argc, line),
        "php.int_max" => {
            crate::emitter::numeric_adapter::emit_php_int_max(chunks, current, argc, line)
        }
        "php.int_min" => {
            crate::emitter::numeric_adapter::emit_php_int_min(chunks, current, argc, line)
        }
        "php.is_int" => {
            crate::emitter::numeric_adapter::emit_php_is_int(chunks, current, argc, line)
        }
        "php.is_float" => {
            crate::emitter::numeric_adapter::emit_php_is_float(chunks, current, argc, line)
        }
        "php.abs" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::NotArray,
                "TypeError",
                "abs(): Argument #1 ($num) must be of type int|float",
                line,
            );
            crate::emitter::numeric_adapter::emit_php_abs(chunks, current, argc, line)
        }
        "php.intdiv" => {
            crate::emitter::numeric_adapter::emit_php_intdiv(chunks, current, argc, line)
        }
        "php.loose_eq" => crate::emitter::relational_adapter::emit_php_loose_eq(
            chunks, current, argc, false, line,
        ),
        "php.loose_ne" => {
            crate::emitter::relational_adapter::emit_php_loose_eq(chunks, current, argc, true, line)
        }
        "php.rand" => crate::emitter::numeric_adapter::emit_rand(chunks, current, argc, line),
        "php.lcg_value" => {
            crate::emitter::numeric_adapter::emit_lcg_value(chunks, current, argc, line)
        }
        "php.pack_float_bytes" => {
            crate::emitter::numeric_adapter::emit_pack_float_bytes(chunks, current, argc, line)
        }

        // ── PHP ctype_* predicates ─────────────────────────────────
        // Char-iteration loops over the input string; each predicate
        // accepts a fixed UTF-16 range set (alpha = A-Z+a-z, etc.).
        "php.ctype_alpha" => {
            crate::emitter::ctype_adapter::emit_ctype_alpha(chunks, current, argc, line)
        }
        "php.ctype_digit" => {
            crate::emitter::ctype_adapter::emit_ctype_digit(chunks, current, argc, line)
        }
        "php.ctype_alnum" => {
            crate::emitter::ctype_adapter::emit_ctype_alnum(chunks, current, argc, line)
        }
        "php.ctype_space" => {
            crate::emitter::ctype_adapter::emit_ctype_space(chunks, current, argc, line)
        }
        "php.ctype_upper" => {
            crate::emitter::ctype_adapter::emit_ctype_upper(chunks, current, argc, line)
        }
        "php.ctype_lower" => {
            crate::emitter::ctype_adapter::emit_ctype_lower(chunks, current, argc, line)
        }
        "php.ctype_xdigit" => {
            crate::emitter::ctype_adapter::emit_ctype_xdigit(chunks, current, argc, line)
        }
        "php.ctype_punct" => {
            crate::emitter::ctype_adapter::emit_ctype_punct(chunks, current, argc, line)
        }
        "php.ctype_print" => {
            crate::emitter::ctype_adapter::emit_ctype_print(chunks, current, argc, line)
        }
        "php.ctype_cntrl" => {
            crate::emitter::ctype_adapter::emit_ctype_cntrl(chunks, current, argc, line)
        }

        // ── PHP math helpers ───────────────────────────────────────
        // min/max (variadic + array form), decbin/decoct/dechex (base
        // string conv), base_convert (string↔string base conv).
        "php.min" => crate::emitter::math_adapter::emit_php_min(chunks, current, argc, line),
        "php.max" => crate::emitter::math_adapter::emit_php_max(chunks, current, argc, line),
        "php.decbin" => crate::emitter::math_adapter::emit_php_decbin(chunks, current, argc, line),
        "php.decoct" => crate::emitter::math_adapter::emit_php_decoct(chunks, current, argc, line),
        "php.dechex" => crate::emitter::math_adapter::emit_php_dechex(chunks, current, argc, line),
        "php.base_convert" => {
            crate::emitter::math_adapter::emit_php_base_convert(chunks, current, argc, line)
        }

        // ── PHP string helpers ─────────────────────────────────────
        // Char/index loops + ECMA string ops (`ecma:string.indexOf`,
        // `STR_TO_LOWER`, `STR_PAD_*`) and `ecma:string.{encode,decode}URIComponent`.
        "php.ucwords" => crate::emitter::string_adapter::emit_ucwords(chunks, current, argc, line),
        "php.strtoupper" => {
            crate::emitter::string_adapter::emit_strtoupper(chunks, current, argc, line)
        }
        "php.addslashes" => {
            crate::emitter::string_adapter::emit_addslashes(chunks, current, argc, line)
        }
        "php.stripslashes" => {
            crate::emitter::string_adapter::emit_stripslashes(chunks, current, argc, line)
        }
        "php.str_rot13" => {
            crate::emitter::string_adapter::emit_str_rot13(chunks, current, argc, line)
        }
        "php.md5" => crate::emitter::string_adapter::emit_md5(chunks, current, argc, line),
        "php.sha1" => crate::emitter::string_adapter::emit_sha1(chunks, current, argc, line),
        "php.hash" => crate::emitter::string_adapter::emit_hash(chunks, current, argc, line),
        "php.hash_hmac" => {
            crate::emitter::string_adapter::emit_hash_hmac(chunks, current, argc, line)
        }
        "php.crc32" => crate::emitter::string_adapter::emit_crc32(chunks, current, argc, line),
        "php.str_split" => {
            crate::emitter::string_adapter::emit_str_split(chunks, current, argc, line)
        }
        "php.base64_decode" => {
            crate::emitter::string_adapter::emit_base64_decode(chunks, current, argc, line)
        }
        "php.explode" => crate::emitter::string_adapter::emit_explode(chunks, current, argc, line),
        "php.sscanf" => vybe_emitter::sprintf::emit_sscanf(chunks, current, argc, line),
        "php.uniqid" => {
            crate::emitter::string_adapter::emit_php_uniqid(chunks, current, argc, line)
        }
        "php.str_pad" => crate::emitter::string_adapter::emit_str_pad(chunks, current, argc, line),
        "php.substr_count" => {
            crate::emitter::string_adapter::emit_substr_count(chunks, current, argc, line)
        }
        "php.substr_replace" => {
            crate::emitter::string_adapter::emit_substr_replace(chunks, current, argc, line)
        }
        "php.str_ireplace" => {
            crate::emitter::string_adapter::emit_str_ireplace(chunks, current, argc, line)
        }
        "php.str_word_count" => {
            crate::emitter::string_adapter::emit_str_word_count(chunks, current, argc, line)
        }
        "php.var_dump_stringify" => {
            crate::emitter::string_adapter::emit_var_dump_stringify(chunks, current, argc, line)
        }
        "php.strstr" => crate::emitter::string_adapter::emit_strstr(chunks, current, argc, line),
        "php.stristr" => crate::emitter::string_adapter::emit_stristr(chunks, current, argc, line),
        "php.strip_tags" => {
            crate::emitter::string_adapter::emit_strip_tags(chunks, current, argc, line)
        }
        "php.strrchr" => crate::emitter::string_adapter::emit_strrchr(chunks, current, argc, line),
        "php.nl2br" => crate::emitter::string_adapter::emit_nl2br(chunks, current, argc, line),
        "php.urlencode" => {
            crate::emitter::string_adapter::emit_urlencode(chunks, current, argc, line)
        }
        "php.rawurlencode" => {
            crate::emitter::string_adapter::emit_rawurlencode(chunks, current, argc, line)
        }
        "php.urldecode" => {
            crate::emitter::string_adapter::emit_urldecode(chunks, current, argc, line)
        }
        "php.rawurldecode" => {
            crate::emitter::string_adapter::emit_rawurldecode(chunks, current, argc, line)
        }
        "php.htmlspecialchars" => {
            crate::emitter::string_adapter::emit_htmlspecialchars(chunks, current, argc, line)
        }
        "php.htmlentities" => {
            crate::emitter::string_adapter::emit_htmlentities(chunks, current, argc, line)
        }
        "php.htmlspecialchars_decode" => {
            crate::emitter::string_adapter::emit_htmlspecialchars_decode(
                chunks, current, argc, line,
            )
        }
        "php.html_entity_decode" => {
            crate::emitter::string_adapter::emit_html_entity_decode(chunks, current, argc, line)
        }
        "php.bin2hex" => crate::emitter::string_adapter::emit_bin2hex(chunks, current, argc, line),
        "php.hex2bin" => crate::emitter::string_adapter::emit_hex2bin(chunks, current, argc, line),
        "php.chunk_split" => {
            crate::emitter::string_adapter::emit_chunk_split(chunks, current, argc, line)
        }
        "php.number_format" => {
            crate::emitter::string_adapter::emit_number_format(chunks, current, argc, line)
        }
        "php.str_replace" => {
            crate::emitter::string_adapter::emit_str_replace(chunks, current, argc, line)
        }
        "php.wordwrap" => {
            crate::emitter::string_adapter::emit_wordwrap(chunks, current, argc, line)
        }
        "php.str_getcsv" => {
            crate::emitter::string_adapter::emit_str_getcsv(chunks, current, argc, line)
        }
        "php.soundex" => crate::emitter::string_adapter::emit_soundex(chunks, current, argc, line),
        "php.levenshtein" => {
            crate::emitter::string_adapter::emit_levenshtein(chunks, current, argc, line)
        }
        "php.similar_text" => {
            crate::emitter::string_adapter::emit_similar_text(chunks, current, argc, line)
        }
        "php.strripos" => {
            crate::emitter::string_adapter::emit_strripos(chunks, current, argc, line)
        }
        "php.strpos" => crate::emitter::string_adapter::emit_strpos(chunks, current, argc, line),
        "php.strtr" => crate::emitter::string_adapter::emit_strtr(chunks, current, argc, line),
        "php.quotemeta" => {
            crate::emitter::string_adapter::emit_quotemeta(chunks, current, argc, line)
        }
        "php.strspn" => crate::emitter::string_adapter::emit_strspn(chunks, current, argc, line),
        "php.strcspn" => crate::emitter::string_adapter::emit_strcspn(chunks, current, argc, line),
        "php.strlen" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::NotArray,
                "TypeError",
                "strlen(): Argument #1 ($string) must be of type string",
                line,
            );
            crate::emitter::string_adapter::emit_strlen(chunks, current, argc, line)
        }
        "php.count_chars" => {
            crate::emitter::string_adapter::emit_count_chars(chunks, current, argc, line)
        }
        "php.convert_uuencode" => {
            crate::emitter::string_adapter::emit_convert_uuencode(chunks, current, argc, line)
        }
        "php.convert_uudecode" => {
            crate::emitter::string_adapter::emit_convert_uudecode(chunks, current, argc, line)
        }
        "php.quoted_printable_encode" => {
            crate::emitter::string_adapter::emit_quoted_printable_encode(
                chunks, current, argc, line,
            )
        }
        "php.quoted_printable_decode" => {
            crate::emitter::string_adapter::emit_quoted_printable_decode(
                chunks, current, argc, line,
            )
        }
        "php.str_increment" => {
            crate::emitter::string_adapter::emit_str_increment(chunks, current, argc, line)
        }
        "php.str_decrement" => {
            crate::emitter::string_adapter::emit_str_decrement(chunks, current, argc, line)
        }
        "php.strncmp" => crate::emitter::string_adapter::emit_strncmp(chunks, current, false, line),
        "php.strncasecmp" => {
            crate::emitter::string_adapter::emit_strncmp(chunks, current, true, line)
        }
        "php.strpbrk" => crate::emitter::string_adapter::emit_strpbrk(chunks, current, argc, line),
        "php.substr_compare" => {
            crate::emitter::string_adapter::emit_substr_compare(chunks, current, argc, line)
        }
        "php.preg_grep" => {
            crate::emitter::string_adapter::emit_preg_grep(chunks, current, argc, line)
        }
        "php.fnmatch" => crate::emitter::string_adapter::emit_fnmatch(chunks, current, argc, line),
        "php.preg_replace_limited" => {
            crate::emitter::string_adapter::emit_preg_replace_limited(chunks, current, argc, line)
        }
        "php.strtok_init" => {
            crate::emitter::string_adapter::emit_strtok_init(chunks, current, argc, line)
        }
        "php.strtok_next" => {
            crate::emitter::string_adapter::emit_strtok_next(chunks, current, argc, line)
        }
        "php.mb_convert_case" => {
            crate::emitter::string_adapter::emit_mb_convert_case(chunks, current, argc, line)
        }
        "php.mb_strwidth" => {
            crate::emitter::string_adapter::emit_mb_strwidth(chunks, current, argc, line)
        }
        "php.mb_language" => crate::emitter::string_adapter::emit_mb_setting(
            chunks,
            current,
            argc,
            "__php_mb_language",
            "neutral",
            line,
        ),
        "php.mb_regex_encoding" => crate::emitter::string_adapter::emit_mb_setting(
            chunks,
            current,
            argc,
            "__php_mb_regex_encoding",
            "UTF-8",
            line,
        ),
        "php.mb_substitute_character" => crate::emitter::string_adapter::emit_mb_setting(
            chunks,
            current,
            argc,
            "__php_mb_substitute_character",
            "none",
            line,
        ),
        "php.mb_check_enc" => {
            let chunk = &mut chunks[current];
            for _ in 0..argc {
                chunk.emit_op(vybe_bytecode::Op::DROP, line);
            }
            chunk.emit_bool_const(true, line);
        }
        "php.mb_detect_enc" => {
            let chunk = &mut chunks[current];
            for _ in 0..argc {
                chunk.emit_op(vybe_bytecode::Op::DROP, line);
            }
            chunk.emit_string_const("UTF-8", line);
        }
        "php.metaphone" => {
            crate::emitter::string_adapter::emit_metaphone(chunks, current, argc, line)
        }
        "php.preg_quote" => {
            crate::emitter::string_adapter::emit_preg_quote(chunks, current, argc, line)
        }
        "php.trim" => crate::emitter::string_adapter::emit_php_trim(chunks, current, argc, line),
        "php.ltrim" => crate::emitter::string_adapter::emit_php_ltrim(chunks, current, argc, line),
        "php.rtrim" => crate::emitter::string_adapter::emit_php_rtrim(chunks, current, argc, line),
        "php.iconv" => crate::emitter::string_adapter::emit_php_iconv(chunks, current, argc, line),
        "php.preg_split" => {
            crate::emitter::string_adapter::emit_preg_split(chunks, current, argc, line)
        }
        "php.preg_match_all_groups" => {
            crate::emitter::string_adapter::emit_preg_match_all_groups(chunks, current, argc, line)
        }
        "php.preg_match_groups" => {
            crate::emitter::string_adapter::emit_preg_match_groups(chunks, current, argc, line)
        }
        "php.preg_replace_callback" => {
            crate::emitter::string_adapter::emit_preg_replace_callback(chunks, current, argc, line)
        }
        "php.clone_helper" => {
            crate::emitter::string_adapter::emit_php_clone(chunks, current, argc, line)
        }
        "php.spl_splstack" => {
            crate::emitter::spl_adapter::emit_spl_new(chunks, current, "SplStack", argc, line)
        }
        "php.spl_splqueue" => {
            crate::emitter::spl_adapter::emit_spl_new(chunks, current, "SplQueue", argc, line)
        }
        "php.spl_spldoublylinkedlist" => crate::emitter::spl_adapter::emit_spl_new(
            chunks,
            current,
            "SplDoublyLinkedList",
            argc,
            line,
        ),
        "php.spl_splminheap" => crate::emitter::spl_adapter::emit_spl_heap_new(
            chunks,
            current,
            "SplMinHeap",
            argc,
            line,
        ),
        "php.spl_splmaxheap" => crate::emitter::spl_adapter::emit_spl_heap_new(
            chunks,
            current,
            "SplMaxHeap",
            argc,
            line,
        ),
        "php.spl_splpriorityqueue" => {
            crate::emitter::spl_adapter::emit_spl_pq_new(chunks, current, argc, line)
        }
        "php.spl_appenditerator" => {
            crate::emitter::spl_adapter::emit_append_iterator_new(chunks, current, argc, line)
        }
        "php.spl_arrayiterator" => {
            crate::emitter::spl_adapter::emit_array_iterator_new(chunks, current, argc, line)
        }
        "php.spl_cachingiterator" => {
            crate::emitter::spl_adapter::emit_caching_iterator_new(chunks, current, argc, line)
        }
        "php.spl_emptyiterator" => {
            crate::emitter::spl_adapter::emit_empty_iterator_new(chunks, current, argc, line)
        }
        "php.spl_infiniteiterator" => {
            crate::emitter::spl_adapter::emit_infinite_iterator_new(chunks, current, argc, line)
        }
        "php.spl_iteratoriterator" => {
            crate::emitter::spl_adapter::emit_iterator_iterator_new(chunks, current, argc, line)
        }
        "php.spl_multipleiterator" => {
            crate::emitter::spl_adapter::emit_multiple_iterator_new(chunks, current, argc, line)
        }
        "php.spl_recursiveiteratoriterator" => {
            crate::emitter::spl_adapter::emit_recursive_iterator_iterator_new(
                chunks, current, argc, line,
            )
        }
        "php.spl_recursivetreeiterator" => {
            crate::emitter::spl_adapter::emit_recursive_tree_iterator_new(
                chunks, current, argc, line,
            )
        }
        "php.spl_fileobject" => {
            crate::emitter::spl_adapter::emit_spl_file_object_new(chunks, current, argc, line)
        }
        "php.spl_tempfileobject" => {
            crate::emitter::spl_adapter::emit_spl_temp_file_object_new(chunks, current, argc, line)
        }
        "php.spl_splobjectstorage" => crate::emitter::spl_adapter::emit_spl_objectstorage_new(
            chunks,
            current,
            "SplObjectStorage",
            argc,
            line,
        ),
        "php.spl_weakmap" => crate::emitter::spl_adapter::emit_spl_objectstorage_new(
            chunks, current, "WeakMap", argc, line,
        ),
        // SplFixedArray is handled by the walker (→ array_fill); no dispatch needed.
        "php.array_merge" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "array_merge(): Argument #1 ($arrays) must be of type array",
                line,
            );
            crate::emitter::array_adapter::emit_php_array_merge(chunks, current, argc, line)
        }
        "php.array_merge_recursive" => {
            crate::emitter::array_adapter::emit_array_merge_recursive(chunks, current, argc, line)
        }
        "php.array_change_key_case" => {
            crate::emitter::array_adapter::emit_array_change_key_case(chunks, current, argc, line)
        }
        "php.array_udiff" => {
            crate::emitter::array_adapter::emit_array_udiff(chunks, current, argc, line)
        }
        "php.array_uintersect" => {
            crate::emitter::array_adapter::emit_array_uintersect(chunks, current, argc, line)
        }
        "php.uniq" => {
            crate::emitter::array_adapter::emit_php_array_unique(chunks, current, argc, line)
        }
        "php.array_union" => {
            crate::emitter::array_adapter::emit_php_array_union(chunks, current, argc, line)
        }
        "php.array_slice" => {
            super::type_guard::guard_arg(
                chunks,
                current,
                argc,
                0,
                super::type_guard::Expect::Array,
                "TypeError",
                "array_slice(): Argument #1 ($array) must be of type array",
                line,
            );
            crate::emitter::array_adapter::emit_php_array_slice(chunks, current, argc, line)
        }
        "php.array_reverse" => {
            crate::emitter::array_adapter::emit_php_array_reverse(chunks, current, argc, line)
        }
        "php.array_rand" => {
            crate::emitter::array_adapter::emit_php_array_rand(chunks, current, argc, line)
        }
        "php.refl_class" => {
            crate::emitter::reflection_adapter::emit_refl_class(chunks, current, argc, line)
        }
        "php.refl_method" => {
            crate::emitter::reflection_adapter::emit_refl_method(chunks, current, argc, line)
        }
        "php.refl_property" => {
            crate::emitter::reflection_adapter::emit_refl_property(chunks, current, argc, line)
        }
        "php.refl_function" => {
            crate::emitter::reflection_adapter::emit_refl_function(chunks, current, argc, line)
        }
        "php.refl_constant" => {
            crate::emitter::reflection_adapter::emit_refl_constant(chunks, current, argc, line)
        }
        "php.weak_ref_create" => {
            crate::emitter::misc_adapter::emit_weak_ref_create(chunks, current, argc, line)
        }
        "php.fiber_new" => {
            crate::emitter::fiber_adapter::emit_php_fiber_new(chunks, current, argc, line)
        }
        "php.fiber_suspend" => {
            crate::emitter::fiber_adapter::emit_php_fiber_suspend(chunks, current, argc, line)
        }
        "php.fiber_start" => {
            crate::emitter::fiber_adapter::emit_php_fiber_start(chunks, current, argc, line)
        }
        "php.fiber_resume" => {
            crate::emitter::fiber_adapter::emit_php_fiber_resume(chunks, current, argc, line)
        }
        "php.fiber_throw" => {
            crate::emitter::fiber_adapter::emit_php_fiber_throw(chunks, current, argc, line)
        }
        "php.fiber_get_return" => {
            crate::emitter::fiber_adapter::emit_php_fiber_get_return(chunks, current, argc, line)
        }
        "php.fiber_is_started" => {
            crate::emitter::fiber_adapter::emit_php_fiber_is_started(chunks, current, argc, line)
        }
        "php.fiber_is_suspended" => {
            crate::emitter::fiber_adapter::emit_php_fiber_is_suspended(chunks, current, argc, line)
        }
        "php.fiber_is_running" => {
            crate::emitter::fiber_adapter::emit_php_fiber_is_running(chunks, current, argc, line)
        }
        "php.fiber_is_terminated" => {
            crate::emitter::fiber_adapter::emit_php_fiber_is_terminated(chunks, current, argc, line)
        }
        "php.echo_stringify" => {
            crate::emitter::string_adapter::emit_echo_stringify(chunks, current, argc, line)
        }

        // ── PHP filesystem helpers ─────────────────────────────────
        // Inline opcode emitters compose the underlying host surfaces
        // into the PHP-facing filesystem API boundary.
        "php.basename" => {
            crate::emitter::filesystem_adapter::emit_basename(chunks, current, argc, line)
        }
        "php.dirname" => {
            crate::emitter::filesystem_adapter::emit_dirname(chunks, current, argc, line)
        }
        "php.file_get_contents" => {
            crate::emitter::filesystem_adapter::emit_file_get_contents(chunks, current, argc, line)
        }
        "php.file_put_contents" => {
            crate::emitter::filesystem_adapter::emit_file_put_contents(chunks, current, argc, line)
        }
        "php.mkdir" => crate::emitter::filesystem_adapter::emit_mkdir(chunks, current, argc, line),
        "php.file_exists" => {
            crate::emitter::filesystem_adapter::emit_file_exists(chunks, current, argc, line)
        }
        "php.is_file" => {
            crate::emitter::filesystem_adapter::emit_is_file(chunks, current, argc, line)
        }
        "php.is_dir" => {
            crate::emitter::filesystem_adapter::emit_is_dir(chunks, current, argc, line)
        }
        "php.is_link" => {
            crate::emitter::filesystem_adapter::emit_is_link(chunks, current, argc, line)
        }
        "php.filesize" => {
            crate::emitter::filesystem_adapter::emit_filesize(chunks, current, argc, line)
        }
        "php.filemtime" => {
            crate::emitter::filesystem_adapter::emit_filemtime(chunks, current, argc, line)
        }
        "php.readlink" => {
            crate::emitter::filesystem_adapter::emit_readlink(chunks, current, argc, line)
        }
        "php.pathinfo" => {
            crate::emitter::filesystem_adapter::emit_pathinfo(chunks, current, argc, line)
        }
        "php.unlink" => {
            crate::emitter::filesystem_adapter::emit_unlink(chunks, current, argc, line)
        }
        "php.file" => crate::emitter::filesystem_adapter::emit_file(chunks, current, argc, line),
        "php.glob" => crate::emitter::filesystem_adapter::emit_glob(chunks, current, argc, line),
        "php.dir" => crate::emitter::filesystem_adapter::emit_dir(chunks, current, argc, line),
        "php.dir_read" => {
            crate::emitter::filesystem_adapter::emit_dir_read(chunks, current, argc, line)
        }
        "php.dir_close" => {
            crate::emitter::filesystem_adapter::emit_dir_close(chunks, current, argc, line)
        }
        "php.sys_get_temp_dir" => {
            crate::emitter::filesystem_adapter::emit_sys_get_temp_dir(chunks, current, argc, line)
        }
        "php.realpath" => {
            crate::emitter::filesystem_adapter::emit_realpath(chunks, current, argc, line)
        }
        "php.copy" => crate::emitter::filesystem_adapter::emit_copy(chunks, current, argc, line),
        "php.rename" => {
            crate::emitter::filesystem_adapter::emit_rename(chunks, current, argc, line)
        }
        "php.rmdir" => crate::emitter::filesystem_adapter::emit_rmdir(chunks, current, argc, line),
        "php.is_readable" => {
            crate::emitter::filesystem_adapter::emit_is_readable(chunks, current, argc, line)
        }
        "php.is_writable" => {
            crate::emitter::filesystem_adapter::emit_is_readable(chunks, current, argc, line)
        }
        "php.filetype" => {
            crate::emitter::filesystem_adapter::emit_filetype(chunks, current, argc, line)
        }
        "php.scandir" => {
            crate::emitter::filesystem_adapter::emit_scandir(chunks, current, argc, line)
        }
        "php.tempnam" => {
            crate::emitter::filesystem_adapter::emit_tempnam(chunks, current, argc, line)
        }
        "php.mime_content_type" => {
            crate::emitter::filesystem_adapter::emit_mime_content_type(chunks, current, argc, line)
        }
        "php.fileperms" => {
            crate::emitter::filesystem_adapter::emit_fileperms(chunks, current, argc, line)
        }
        "php.disk_free_space" => {
            crate::emitter::filesystem_adapter::emit_disk_free_space(chunks, current, argc, line)
        }
        "php.disk_total_space" => {
            crate::emitter::filesystem_adapter::emit_disk_total_space(chunks, current, argc, line)
        }
        "php.fopen" => crate::emitter::filesystem_adapter::emit_fopen(chunks, current, argc, line),
        "php.fwrite" => {
            crate::emitter::filesystem_adapter::emit_fwrite(chunks, current, argc, line)
        }
        "php.fread" => crate::emitter::filesystem_adapter::emit_fread(chunks, current, argc, line),
        "php.fgets" => crate::emitter::filesystem_adapter::emit_fgets(chunks, current, argc, line),
        "php.fgetc" => crate::emitter::filesystem_adapter::emit_fgetc(chunks, current, argc, line),
        "php.feof" => crate::emitter::filesystem_adapter::emit_feof(chunks, current, argc, line),
        "php.ftell" => crate::emitter::filesystem_adapter::emit_ftell(chunks, current, argc, line),
        "php.fseek" => crate::emitter::filesystem_adapter::emit_fseek(chunks, current, argc, line),
        "php.rewind" => {
            crate::emitter::filesystem_adapter::emit_rewind(chunks, current, argc, line)
        }
        "php.fflush" => {
            crate::emitter::filesystem_adapter::emit_fflush(chunks, current, argc, line)
        }
        "php.fclose" => {
            crate::emitter::filesystem_adapter::emit_fclose(chunks, current, argc, line)
        }
        "php.stream_get_contents" => crate::emitter::filesystem_adapter::emit_stream_get_contents(
            chunks, current, argc, line,
        ),
        "php.stream_get_meta_data" => {
            crate::emitter::filesystem_adapter::emit_stream_get_meta_data(
                chunks, current, argc, line,
            )
        }

        // ── .NET String.Format adapter — composite-format substitution ──
        // `String.Format(fmt, ...args)` lowers to inline bytecode that
        // walks the format string, parses `{N}` / `{{` / `}}` tokens,
        // and concatenates. `argc` includes the format string; trailing
        // args are packed into a local array indexed by placeholder N.
        "php.isset_all" => emit_isset_all(&mut chunks[current], argc, line),
        "php.header" => crate::emitter::misc_adapter::emit_php_header(chunks, current, argc, line),
        "php.extension_loaded" => {
            crate::emitter::misc_adapter::emit_php_extension_loaded(chunks, current, argc, line)
        }
        "php.phpversion" => {
            crate::emitter::misc_adapter::emit_php_phpversion(chunks, current, argc, line)
        }
        "php.phpinfo" => {
            crate::emitter::misc_adapter::emit_php_phpinfo(chunks, current, argc, line)
        }
        "php.spl_autoload_register" => {
            crate::emitter::misc_adapter::emit_php_spl_autoload_register(
                chunks, current, argc, line,
            )
        }
        "php.spl_autoload_unregister" => {
            crate::emitter::misc_adapter::emit_php_spl_autoload_unregister(
                chunks, current, argc, line,
            )
        }
        "php.empty" => crate::emitter::misc_adapter::emit_php_empty(chunks, current, argc, line),
        "php.session_start" => {
            crate::emitter::misc_adapter::emit_php_session_start(chunks, current, argc, line)
        }
        "php.session_unset" => {
            crate::emitter::misc_adapter::emit_php_session_unset(chunks, current, argc, line)
        }
        "php.session_destroy" => {
            crate::emitter::misc_adapter::emit_php_session_destroy(chunks, current, argc, line)
        }
        "php.serialize" => {
            crate::emitter::misc_adapter::emit_php_serialize(chunks, current, argc, line)
        }
        "php.unserialize" => {
            crate::emitter::misc_adapter::emit_php_unserialize(chunks, current, argc, line)
        }
        "php.pdo_new" => crate::emitter::pdo_adapter::emit_php_pdo_new(chunks, current, argc, line),
        "php.pdo_query" => {
            crate::emitter::pdo_adapter::emit_php_pdo_query(chunks, current, argc, line)
        }
        "php.pdo_exec" => {
            crate::emitter::pdo_adapter::emit_php_pdo_exec(chunks, current, argc, line)
        }
        "php.pdo_prepare" => {
            crate::emitter::pdo_adapter::emit_php_pdo_prepare(chunks, current, argc, line)
        }
        "php.db_prepare" => {
            crate::emitter::db_adapter::emit_db_prepare(chunks, current, argc, line)
        }
        "php.set_error_handler" => {
            super::error_adapter::emit_set_error_handler(chunks, current, argc, line)
        }
        "php.assert_active" => {
            super::error_adapter::emit_assert_active(chunks, current, argc, line)
        }
        "php.assert_active_option" => {
            super::error_adapter::emit_assert_active_option(chunks, current, argc, line)
        }
        "php.assert_callback" => {
            super::error_adapter::emit_assert_callback(chunks, current, argc, line)
        }
        "php.assert_callback_option" => {
            super::error_adapter::emit_assert_callback_option(chunks, current, argc, line)
        }
        "php.restore_error_handler" => {
            super::error_adapter::emit_restore_error_handler(chunks, current, argc, line)
        }
        "php.trigger_error" => {
            super::error_adapter::emit_trigger_error(chunks, current, argc, line)
        }
        "php.error_get_last" => {
            super::error_adapter::emit_error_get_last(chunks, current, argc, line)
        }
        "php.error_clear_last" => {
            super::error_adapter::emit_error_clear_last(chunks, current, argc, line)
        }
        "php.ob_start" => super::output_adapter::emit_ob_start(chunks, current, argc, line),
        "php.ob_get_clean" => super::output_adapter::emit_ob_get_clean(chunks, current, argc, line),
        "php.ob_get_contents" => {
            super::output_adapter::emit_ob_get_contents(chunks, current, argc, line)
        }
        "php.ob_end_clean" => super::output_adapter::emit_ob_end_clean(chunks, current, argc, line),
        "php.ob_end_flush" => super::output_adapter::emit_ob_end_flush(chunks, current, argc, line),
        "php.ob_get_level" => super::output_adapter::emit_ob_get_level(chunks, current, argc, line),
        "php.ob_clean" => super::output_adapter::emit_ob_clean(chunks, current, argc, line),
        "php.ob_get_length" => {
            super::output_adapter::emit_ob_get_length(chunks, current, argc, line)
        }
        "php.ob_implicit_flush" => {
            super::output_adapter::emit_ob_implicit_flush(chunks, current, argc, line)
        }
        "php.ob_list_handlers" => {
            super::output_adapter::emit_ob_list_handlers(chunks, current, argc, line)
        }
        "php.ob_gzhandler" => super::output_adapter::emit_ob_gzhandler(chunks, current, argc, line),
        "php.error_reporting" => {
            super::error_adapter::emit_error_reporting(chunks, current, argc, line)
        }
        "php.set_exception_handler" => {
            super::error_adapter::emit_set_exception_handler(chunks, current, argc, line)
        }
        "php.pdo_statement_fetch_column" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_fetch_column(
                chunks, current, argc, line,
            )
        }
        "php.pdo_statement_row_count" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_row_count(
                chunks, current, argc, line,
            )
        }
        "php.pdo_statement_column_count" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_column_count(
                chunks, current, argc, line,
            )
        }
        "php.pdo_get_attribute" => {
            crate::emitter::pdo_adapter::emit_php_pdo_get_attribute(chunks, current, argc, line)
        }
        "php.pdo_quote" => {
            crate::emitter::pdo_adapter::emit_php_pdo_quote(chunks, current, argc, line)
        }
        "php.pdo_error_code" => {
            crate::emitter::pdo_adapter::emit_php_pdo_error_code(chunks, current, argc, line)
        }
        "php.pdo_last_insert_id" => {
            crate::emitter::pdo_adapter::emit_php_pdo_last_insert_id(chunks, current, argc, line)
        }
        "php.pdo_error_info" => {
            crate::emitter::pdo_adapter::emit_php_pdo_error_info(chunks, current, argc, line)
        }
        "php.mysqli_stmt_attr_get" => {
            super::mysqli_adapter::emit_mysqli_stmt_attr_get(chunks, current, argc, line)
        }
        "php.mysqli_stmt_true" => {
            super::mysqli_adapter::emit_mysqli_stmt_true(chunks, current, argc, line)
        }
        "php.mysqli_stmt_false" => {
            super::mysqli_adapter::emit_mysqli_stmt_false(chunks, current, argc, line)
        }
        "php.mysqli_stmt_get_result" => {
            super::mysqli_adapter::emit_php_mysqli_stmt_get_result(chunks, current, argc, line)
        }
        "php.pdo_set_attribute" => {
            crate::emitter::pdo_adapter::emit_php_pdo_set_attribute(chunks, current, argc, line)
        }
        "php.pdo_begin_transaction" => {
            crate::emitter::pdo_adapter::emit_php_pdo_begin_transaction(chunks, current, argc, line)
        }
        "php.pdo_commit" => {
            crate::emitter::pdo_adapter::emit_php_pdo_commit(chunks, current, argc, line)
        }
        "php.pdo_rollback" => {
            crate::emitter::pdo_adapter::emit_php_pdo_rollback(chunks, current, argc, line)
        }
        "php.pdo_statement_bind_param" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_bind_param(
                chunks, current, argc, line,
            )
        }
        "php.pdo_statement_bind_value" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_bind_value(
                chunks, current, argc, line,
            )
        }
        "php.pdo_statement_bind_column" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_bind_column(
                chunks, current, argc, line,
            )
        }
        "php.pdo_statement_execute" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_execute(chunks, current, argc, line)
        }
        "php.pdo_statement_fetch" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_fetch(chunks, current, argc, line)
        }
        "php.pdo_statement_fetch_all" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_fetch_all(
                chunks, current, argc, line,
            )
        }
        "php.pdo_statement_param_count" => {
            crate::emitter::pdo_adapter::emit_php_pdo_statement_param_count(
                chunks, current, argc, line,
            )
        }
        "php.mysqli_report" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_report(chunks, current, argc, line)
        }
        "php.mysqli_connect" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_connect(chunks, current, argc, line)
        }
        "php.mysqli_init" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_init(chunks, current, argc, line)
        }
        "php.mysqli_real_connect" => crate::emitter::mysqli_adapter::emit_php_mysqli_real_connect(
            chunks, current, argc, line,
        ),
        "php.mysqli_connect_errno" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_connect_errno(
                chunks, current, argc, line,
            )
        }
        "php.mysqli_connect_error" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_connect_error(
                chunks, current, argc, line,
            )
        }
        "php.mysqli_error" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_error(chunks, current, argc, line)
        }
        "php.mysqli_query" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_query(chunks, current, argc, line)
        }
        "php.mysqli_prepare" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_prepare(chunks, current, argc, line)
        }
        "php.mysqli_select_db" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_select_db(chunks, current, argc, line)
        }
        "php.mysqli_set_charset" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_set_charset(chunks, current, argc, line)
        }
        "php.mysqli_ping" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_ping(chunks, current, argc, line)
        }
        "php.mysqli_errno" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_errno(chunks, current, argc, line)
        }
        "php.mysqli_affected_rows" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_affected_rows(
                chunks, current, argc, line,
            )
        }
        "php.mysqli_insert_id" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_insert_id(chunks, current, argc, line)
        }
        "php.mysqli_num_fields" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_num_fields(chunks, current, argc, line)
        }
        "php.mysqli_fetch_field" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_fetch_field(chunks, current, argc, line)
        }
        "php.mysqli_free_result" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_free_result(chunks, current, argc, line)
        }
        "php.mysqli_more_results" => crate::emitter::mysqli_adapter::emit_php_mysqli_more_results(
            chunks, current, argc, line,
        ),
        "php.mysqli_next_result" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_next_result(chunks, current, argc, line)
        }
        "php.mysqli_close" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_close(chunks, current, argc, line)
        }
        "php.mysqli_real_escape_string" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_real_escape_string(
                chunks, current, argc, line,
            )
        }
        "php.mysqli_character_set_name" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_character_set_name(
                chunks, current, argc, line,
            )
        }
        "php.mysqli_get_client_info" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_get_client_info(
                chunks, current, argc, line,
            )
        }
        "php.mysqli_get_server_info" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_get_server_info(
                chunks, current, argc, line,
            )
        }
        "php.mysqli_fetch_array" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_fetch_array(chunks, current, argc, line)
        }
        "php.mysqli_fetch_assoc" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_fetch_assoc(chunks, current, argc, line)
        }
        "php.mysqli_fetch_object" => crate::emitter::mysqli_adapter::emit_php_mysqli_fetch_object(
            chunks, current, argc, line,
        ),
        "php.mysqli_num_rows" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_num_rows(chunks, current, argc, line)
        }
        "php.mysqli_fetch_all" => {
            crate::emitter::mysqli_adapter::emit_php_mysqli_fetch_all(chunks, current, argc, line)
        }

        // ── Fortran `max(a, b, c, ...)` / `min(a, b, c, ...)` — variadic.
        // Pure WASM (chained f64.max / f64.min); no host calls.
        _ => return false,
    }
    true
}
