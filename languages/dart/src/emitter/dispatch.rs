//! Auto-extracted `dart.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "dart.is_empty" => {
            crate::emitter::string_adapter::emit_dart_is_empty(chunks, current, line)
        }
        "dart.is_not_empty" => {
            crate::emitter::string_adapter::emit_dart_is_not_empty(chunks, current, line)
        }
        "dart.is_even" => {
            crate::emitter::string_adapter::emit_dart_is_even(chunks, current, line)
        }
        "dart.is_odd" => {
            crate::emitter::string_adapter::emit_dart_is_odd(chunks, current, line)
        }
        "dart.sb_new" => {
            crate::emitter::string_adapter::emit_dart_sb_new(chunks, current, line)
        }
        "dart.sb_write" => {
            crate::emitter::string_adapter::emit_dart_sb_write(chunks, current, line)
        }
        "dart.sb_writeln" => {
            crate::emitter::string_adapter::emit_dart_sb_writeln(chunks, current, argc, line)
        }
        "dart.sb_write_all" => {
            crate::emitter::string_adapter::emit_dart_sb_write_all(chunks, current, argc, line)
        }
        "dart.sb_write_char_code" => {
            crate::emitter::string_adapter::emit_dart_sb_write_char_code(chunks, current, line)
        }
        "dart.sb_clear" => {
            crate::emitter::string_adapter::emit_dart_sb_clear(chunks, current, line)
        }
        "dart.regexp_new" => {
            crate::emitter::string_adapter::emit_dart_regexp_new(chunks, current, argc, line)
        }
        "dart.regexp_has_match" => {
            crate::emitter::string_adapter::emit_dart_regexp_has_match(chunks, current, line)
        }
        "dart.regexp_first_match" => {
            crate::emitter::string_adapter::emit_dart_regexp_first_match(chunks, current, line)
        }
        "dart.regexp_all_matches" => {
            crate::emitter::string_adapter::emit_dart_regexp_all_matches(chunks, current, line)
        }
        "dart.regexp_group" => {
            crate::emitter::string_adapter::emit_dart_regexp_group(chunks, current, line)
        }
        "dart.duration_new" => crate::emitter::core_adapter::emit_duration_new(chunks, current, argc, line),
        "dart.duration_zero" => crate::emitter::core_adapter::emit_duration_zero(chunks, current, line),
        "dart.duration_abs" => crate::emitter::core_adapter::emit_duration_abs(chunks, current, line),
        "dart.duration_negate" => crate::emitter::core_adapter::emit_duration_negate(chunks, current, line),
        "dart.datetime_new" => crate::emitter::core_adapter::emit_datetime_new(chunks, current, argc, false, line),
        "dart.datetime_utc" => crate::emitter::core_adapter::emit_datetime_new(chunks, current, argc, true, line),
        "dart.add" => crate::emitter::core_adapter::emit_dart_add(chunks, current, line),
        "dart.datetime_add" => crate::emitter::core_adapter::emit_datetime_add(chunks, current, line),
        "dart.datetime_subtract" => crate::emitter::core_adapter::emit_datetime_subtract(chunks, current, line),
        "dart.datetime_difference" => crate::emitter::core_adapter::emit_datetime_difference(chunks, current, line),
        "dart.datetime_is_before" => crate::emitter::core_adapter::emit_datetime_is_before(chunks, current, line),
        "dart.datetime_is_after" => crate::emitter::core_adapter::emit_datetime_is_after(chunks, current, line),
        "dart.datetime_same_moment" => crate::emitter::core_adapter::emit_datetime_same_moment(chunks, current, line),
        "dart.compare_to" => crate::emitter::core_adapter::emit_compare_to(chunks, current, line),
        "dart.uri_parse" => crate::emitter::core_adapter::emit_uri_parse(chunks, current, line),
        "dart.uri_http" => crate::emitter::core_adapter::emit_uri_http(chunks, current, argc, false, line),
        "dart.uri_https" => crate::emitter::core_adapter::emit_uri_http(chunks, current, argc, true, line),
        "dart.uri_file" => crate::emitter::core_adapter::emit_uri_file(chunks, current, line),
        "dart.uri_normalize_path" => crate::emitter::core_adapter::emit_uri_normalize_path(chunks, current, line),
        "dart.uri_replace" => crate::emitter::core_adapter::emit_uri_replace(chunks, current, line),
        "dart.uri_resolve" => crate::emitter::core_adapter::emit_uri_resolve(chunks, current, line),
        "dart.uri_resolve_uri" => crate::emitter::core_adapter::emit_uri_resolve_uri(chunks, current, line),
        "dart.list_filled" => crate::emitter::core_adapter::emit_list_filled(chunks, current, line),
        "dart.replace_first" => {
            crate::emitter::string_adapter::emit_dart_replace_first(chunks, current, line)
        }
        "dart.list_first" => {
            crate::emitter::string_adapter::emit_dart_list_first(chunks, current, line)
        }
        "dart.list_last" => {
            crate::emitter::string_adapter::emit_dart_list_last(chunks, current, line)
        }
        "dart.length" => {
            crate::emitter::string_adapter::emit_dart_length(chunks, current, line)
        }
        "dart.print" => {
            crate::emitter::string_adapter::emit_dart_print(chunks, current, argc, line)
        }
        "dart.to_string" => {
            crate::emitter::string_adapter::emit_dart_to_string(chunks, current, line)
        }

        // ── Ruby `obj.dig(k1, k2, ..., kN)` — variadic property walk.
        // Returns `obj[k1]?[k2]?...[kN]`, or `nil` if any link is null.
        // `argc` includes receiver: `argc == N + 1`. Inline emit chains
        // ARRAY_GET (polymorphic over Map / Object / Array) with
        // null-short-circuit at every step.
        _ => return false,
    }
    true
}
