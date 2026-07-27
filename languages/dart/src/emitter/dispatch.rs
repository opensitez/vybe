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
        "dart.is_even" => crate::emitter::string_adapter::emit_dart_is_even(chunks, current, line),
        "dart.is_odd" => crate::emitter::string_adapter::emit_dart_is_odd(chunks, current, line),
        "dart.sb_new" => crate::emitter::string_adapter::emit_dart_sb_new(chunks, current, line),
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
        "dart.exception" => {
            crate::emitter::string_adapter::emit_dart_exception(chunks, current, argc, line)
        }
        "dart.format_exception" => {
            crate::emitter::string_adapter::emit_dart_format_exception(chunks, current, argc, line)
        }
        "dart.range_error" => {
            crate::emitter::string_adapter::emit_dart_range_error(chunks, current, argc, line)
        }
        "dart.state_error" => {
            crate::emitter::string_adapter::emit_dart_state_error(chunks, current, argc, line)
        }
        "dart.argument_error" => {
            crate::emitter::string_adapter::emit_dart_argument_error(chunks, current, argc, line)
        }
        "dart.unimplemented_error" => {
            crate::emitter::string_adapter::emit_dart_unimplemented_error(
                chunks, current, argc, line,
            )
        }
        "dart.stack_trace" => {
            crate::emitter::string_adapter::emit_dart_stack_trace(chunks, current, line)
        }
        "dart.index_get" => {
            crate::emitter::string_adapter::emit_dart_index_get(chunks, current, line)
        }
        "dart.eq" => crate::emitter::string_adapter::emit_dart_eq(chunks, current, line),
        "dart.is_null" => crate::emitter::string_adapter::emit_dart_is_null(chunks, current, line),
        "dart.identical" => {
            crate::emitter::string_adapter::emit_dart_identical(chunks, current, line)
        }
        "dart.hash_code" => {
            crate::emitter::string_adapter::emit_dart_hash_code(chunks, current, line)
        }
        "dart.object_hash_code" => {
            crate::emitter::string_adapter::emit_dart_object_hash_code(chunks, current, line)
        }
        "dart.runtime_type" => {
            crate::emitter::reflection_adapter::emit_dart_runtime_type(chunks, current, line)
        }
        "dart.type_to_string" => {
            crate::emitter::reflection_adapter::emit_dart_type_to_string(chunks, current, line)
        }
        "dart.is_list_of_int" => {
            crate::emitter::reflection_adapter::emit_dart_is_list_of_int(chunks, current, line)
        }
        "dart.duration_new" => {
            crate::emitter::core_adapter::emit_duration_new(chunks, current, argc, line)
        }
        "dart.duration_zero" => {
            crate::emitter::core_adapter::emit_duration_zero(chunks, current, line)
        }
        "dart.abs" => crate::emitter::core_adapter::emit_dart_abs(chunks, current, line),
        "dart.num_floor" => crate::emitter::core_adapter::emit_num_floor(chunks, current, line),
        "dart.num_ceil" => crate::emitter::core_adapter::emit_num_ceil(chunks, current, line),
        "dart.num_round" => crate::emitter::core_adapter::emit_num_round(chunks, current, line),
        "dart.num_truncate" => {
            crate::emitter::core_adapter::emit_num_truncate(chunks, current, line)
        }
        "dart.num_to_double" => {
            crate::emitter::core_adapter::emit_num_to_double(chunks, current, line)
        }
        "dart.num_remainder" => {
            crate::emitter::core_adapter::emit_num_remainder(chunks, current, line)
        }
        "dart.num_is_negative" => {
            crate::emitter::core_adapter::emit_num_is_negative(chunks, current, line)
        }
        "dart.num_is_infinite" => {
            crate::emitter::core_adapter::emit_num_is_infinite(chunks, current, line)
        }
        "dart.num_sign" => crate::emitter::core_adapter::emit_num_sign(chunks, current, line),
        "dart.duration_negate" => {
            crate::emitter::core_adapter::emit_duration_negate(chunks, current, line)
        }
        "dart.datetime_new" => {
            crate::emitter::core_adapter::emit_datetime_new(chunks, current, argc, false, line)
        }
        "dart.datetime_utc" => {
            crate::emitter::core_adapter::emit_datetime_new(chunks, current, argc, true, line)
        }
        "dart.add_general" => {
            crate::emitter::string_adapter::emit_dart_add_general(chunks, current, argc, line)
        }
        "dart.index_set" => {
            crate::emitter::string_adapter::emit_dart_index_set(chunks, current, line)
        }
        "dart.difference" => {
            crate::emitter::string_adapter::emit_dart_difference(chunks, current, line)
        }
        "dart.datetime_add" => {
            crate::emitter::core_adapter::emit_datetime_add(chunks, current, line)
        }
        "dart.datetime_subtract" => {
            crate::emitter::core_adapter::emit_datetime_subtract(chunks, current, line)
        }
        "dart.datetime_difference" => {
            crate::emitter::core_adapter::emit_datetime_difference(chunks, current, line)
        }
        "dart.datetime_is_before" => {
            crate::emitter::core_adapter::emit_datetime_is_before(chunks, current, line)
        }
        "dart.datetime_is_after" => {
            crate::emitter::core_adapter::emit_datetime_is_after(chunks, current, line)
        }
        "dart.datetime_same_moment" => {
            crate::emitter::core_adapter::emit_datetime_same_moment(chunks, current, line)
        }
        "dart.compare_to" => crate::emitter::core_adapter::emit_compare_to(chunks, current, line),
        "dart.uri_parse" => crate::emitter::core_adapter::emit_uri_parse(chunks, current, line),
        "dart.uri_http" => {
            crate::emitter::core_adapter::emit_uri_http(chunks, current, argc, false, line)
        }
        "dart.uri_https" => {
            crate::emitter::core_adapter::emit_uri_http(chunks, current, argc, true, line)
        }
        "dart.uri_file" => crate::emitter::core_adapter::emit_uri_file(chunks, current, line),
        "dart.uri_normalize_path" => {
            crate::emitter::core_adapter::emit_uri_normalize_path(chunks, current, line)
        }
        "dart.uri_replace" => crate::emitter::core_adapter::emit_uri_replace(chunks, current, line),
        "dart.uri_resolve" => crate::emitter::core_adapter::emit_uri_resolve(chunks, current, line),
        "dart.uri_resolve_uri" => {
            crate::emitter::core_adapter::emit_uri_resolve_uri(chunks, current, line)
        }
        "dart.list_filled" => crate::emitter::core_adapter::emit_list_filled(chunks, current, line),
        "dart.list_generate" => {
            crate::emitter::string_adapter::emit_dart_list_generate(chunks, current, argc, line)
        }
        "dart.list_from" => {
            crate::emitter::string_adapter::emit_dart_list_from(chunks, current, argc, line)
        }
        "dart.list_unmodifiable" => {
            crate::emitter::string_adapter::emit_dart_list_unmodifiable(chunks, current, line)
        }
        "dart.map_new" => crate::emitter::string_adapter::emit_dart_map_new(chunks, current, line),
        "dart.sorted_map_new" => {
            crate::emitter::string_adapter::emit_dart_sorted_map_new(chunks, current, line)
        }
        "dart.set_new" => crate::emitter::string_adapter::emit_dart_set_new(chunks, current, line),
        // SplayTreeSet.add — dedupe + insert + keep ascending via the shared
        // sorted core (same engine as Java TreeSet / .NET SortedSet). The
        // tagged-array set backing is unchanged; only `.add` sorts.
        "dart.sorted_set_add" => {
            vybe_compiler::compiler::sorted_collection::emit_sorted_add(chunks, current, line)
        }
        "dart.set_from" => {
            crate::emitter::string_adapter::emit_dart_set_from(chunks, current, line)
        }
        "dart.set_unmodifiable" => {
            crate::emitter::string_adapter::emit_dart_set_unmodifiable(chunks, current, line)
        }
        "dart.map_entry" => {
            crate::emitter::string_adapter::emit_dart_map_entry(chunks, current, line)
        }
        "dart.map_from" => {
            crate::emitter::string_adapter::emit_dart_map_from(chunks, current, line)
        }
        "dart.map_unmodifiable" => {
            crate::emitter::string_adapter::emit_dart_map_unmodifiable(chunks, current, line)
        }
        "dart.map_unmodifiable_entries" => {
            crate::emitter::string_adapter::emit_dart_map_unmodifiable_entries(
                chunks, current, line,
            )
        }
        "dart.map_from_entries" => {
            crate::emitter::string_adapter::emit_dart_map_from_entries(chunks, current, line)
        }
        "dart.map_from_iterables" => {
            crate::emitter::string_adapter::emit_dart_map_from_iterables(chunks, current, line)
        }
        "dart.identity" => {
            crate::emitter::string_adapter::emit_dart_identity(chunks, current, line)
        }
        "dart.string_from_char_codes" => {
            crate::emitter::string_adapter::emit_dart_string_from_char_codes(chunks, current, line)
        }
        "dart.string_code_units" => {
            crate::emitter::string_adapter::emit_dart_string_code_units(chunks, current, line)
        }
        "dart.string_runes" => {
            crate::emitter::string_adapter::emit_dart_string_runes(chunks, current, line)
        }
        "dart.replace_first" => {
            crate::emitter::string_adapter::emit_dart_replace_first(chunks, current, line)
        }
        "dart.list_first" => {
            crate::emitter::string_adapter::emit_dart_list_first(chunks, current, line)
        }
        "dart.list_last" => {
            crate::emitter::string_adapter::emit_dart_list_last(chunks, current, line)
        }
        "dart.list_single" => {
            crate::emitter::string_adapter::emit_dart_list_single(chunks, current, false, line)
        }
        "dart.list_single_or_null" => {
            crate::emitter::string_adapter::emit_dart_list_single(chunks, current, true, line)
        }
        "dart.length" => crate::emitter::string_adapter::emit_dart_length(chunks, current, line),
        "dart.map_keys" => {
            crate::emitter::string_adapter::emit_dart_map_keys(chunks, current, line)
        }
        "dart.map_values" => {
            crate::emitter::string_adapter::emit_dart_map_values(chunks, current, line)
        }
        "dart.map_entries" => {
            crate::emitter::string_adapter::emit_dart_map_entries(chunks, current, line)
        }
        "dart.map_contains_value" => {
            crate::emitter::string_adapter::emit_dart_map_contains_value(chunks, current, line)
        }
        "dart.add_all" => crate::emitter::string_adapter::emit_dart_add_all(chunks, current, line),
        "dart.index_of" => {
            crate::emitter::string_adapter::emit_dart_index_of(chunks, current, argc, false, line)
        }
        "dart.last_index_of" => {
            crate::emitter::string_adapter::emit_dart_index_of(chunks, current, argc, true, line)
        }
        "dart.list_insert" => {
            crate::emitter::string_adapter::emit_dart_list_insert(chunks, current, line)
        }
        "dart.list_remove_at" => {
            crate::emitter::string_adapter::emit_dart_list_remove_at(chunks, current, line)
        }
        "dart.list_remove_last" => {
            crate::emitter::string_adapter::emit_dart_list_remove_last(chunks, current, line)
        }
        "dart.list_remove_range" => {
            crate::emitter::string_adapter::emit_dart_list_remove_range(chunks, current, line)
        }
        "dart.list_get_range" => {
            crate::emitter::string_adapter::emit_dart_list_get_range(chunks, current, line)
        }
        "dart.list_fill_range" => {
            crate::emitter::string_adapter::emit_dart_list_fill_range(chunks, current, line)
        }
        "dart.list_set_all" => {
            crate::emitter::string_adapter::emit_dart_list_set_all(chunks, current, line)
        }
        "dart.list_set_range" => {
            crate::emitter::string_adapter::emit_dart_list_set_range(chunks, current, line)
        }
        "dart.list_as_map" => {
            crate::emitter::string_adapter::emit_dart_list_as_map(chunks, current, line)
        }
        "dart.list_sort" => {
            crate::emitter::string_adapter::emit_dart_list_sort(chunks, current, argc, line)
        }
        "dart.list_reversed" => {
            crate::emitter::string_adapter::emit_dart_list_reversed(chunks, current, line)
        }
        "dart.remove" => crate::emitter::string_adapter::emit_dart_remove(chunks, current, line),
        "dart.clear" => crate::emitter::string_adapter::emit_dart_clear(chunks, current, line),
        "dart.lookup" => crate::emitter::string_adapter::emit_dart_lookup(chunks, current, line),
        "dart.remove_where" => {
            crate::emitter::string_adapter::emit_dart_remove_where(chunks, current, line)
        }
        "dart.set_remove_all" => {
            crate::emitter::string_adapter::emit_dart_set_remove_all(chunks, current, line)
        }
        "dart.set_retain_all" => {
            crate::emitter::string_adapter::emit_dart_set_retain_all(chunks, current, line)
        }
        "dart.set_retain_where" => {
            crate::emitter::string_adapter::emit_dart_set_retain_where(chunks, current, line)
        }
        "dart.set_union" => {
            crate::emitter::string_adapter::emit_dart_set_union(chunks, current, line)
        }
        "dart.set_intersection" => {
            crate::emitter::string_adapter::emit_dart_set_intersection(chunks, current, line)
        }
        "dart.set_contains_all" => {
            crate::emitter::string_adapter::emit_dart_set_contains_all(chunks, current, line)
        }
        "dart.map_update" => {
            crate::emitter::string_adapter::emit_dart_map_update(chunks, current, argc, line)
        }
        "dart.map_put_if_absent" => {
            crate::emitter::string_adapter::emit_dart_map_put_if_absent(chunks, current, line)
        }
        "dart.map_update_all" => {
            crate::emitter::string_adapter::emit_dart_map_update_all(chunks, current, line)
        }
        "dart.iter_to_list" => {
            crate::emitter::string_adapter::emit_dart_iter_to_list(chunks, current, line)
        }
        "dart.for_in_iterable" => {
            crate::emitter::string_adapter::emit_dart_for_in_iterable(chunks, current, line)
        }
        "dart.iter_join" => {
            crate::emitter::string_adapter::emit_dart_iter_join(chunks, current, argc, line)
        }
        "dart.int_parse" => {
            crate::emitter::string_adapter::emit_dart_int_parse(chunks, current, argc, line)
        }
        "dart.int_try_parse" => {
            crate::emitter::string_adapter::emit_dart_int_try_parse(chunks, current, argc, line)
        }
        "dart.double_parse" => {
            crate::emitter::string_adapter::emit_dart_double_parse(chunks, current, line)
        }
        "dart.double_try_parse" => {
            crate::emitter::string_adapter::emit_dart_double_try_parse(chunks, current, line)
        }
        "dart.bigint_from" => {
            crate::emitter::string_adapter::emit_dart_bigint_from(chunks, current, line)
        }
        "dart.stream_value" => {
            crate::emitter::string_adapter::emit_dart_stream_value(chunks, current, line)
        }
        "dart.stream_empty" => {
            crate::emitter::string_adapter::emit_dart_stream_empty(chunks, current, line)
        }
        "dart.stream_error" => {
            crate::emitter::string_adapter::emit_dart_stream_error(chunks, current, line)
        }
        "dart.stream_listen" => {
            crate::emitter::string_adapter::emit_dart_stream_listen(chunks, current, argc, line)
        }
        "dart.stream_as_future" => {
            crate::emitter::string_adapter::emit_dart_stream_as_future(chunks, current, line)
        }
        "dart.stream_cancel" => {
            crate::emitter::string_adapter::emit_dart_stream_cancel(chunks, current, line)
        }
        "dart.queue_remove_first" => {
            crate::emitter::string_adapter::emit_dart_queue_remove_first(chunks, current, line)
        }
        "dart.bigint_parse" => {
            crate::emitter::string_adapter::emit_dart_bigint_parse(chunks, current, argc, line)
        }
        "dart.bigint_gcd" => {
            crate::emitter::string_adapter::emit_dart_bigint_gcd(chunks, current, line)
        }
        "dart.stopwatch_new" => {
            crate::emitter::string_adapter::emit_dart_stopwatch_new(chunks, current, line)
        }
        "dart.stopwatch_start" => {
            crate::emitter::string_adapter::emit_dart_stopwatch_start(chunks, current, line)
        }
        "dart.stopwatch_stop" => {
            crate::emitter::string_adapter::emit_dart_stopwatch_stop(chunks, current, line)
        }
        "dart.stopwatch_reset" => {
            crate::emitter::string_adapter::emit_dart_stopwatch_reset(chunks, current, line)
        }
        "dart.stopwatch_is_running" => {
            crate::emitter::string_adapter::emit_dart_stopwatch_is_running(chunks, current, line)
        }
        "dart.stopwatch_elapsed" => {
            crate::emitter::string_adapter::emit_dart_stopwatch_elapsed(chunks, current, line)
        }
        "dart.stopwatch_elapsed_milliseconds" => {
            crate::emitter::string_adapter::emit_dart_stopwatch_elapsed_milliseconds(
                chunks, current, line,
            )
        }
        "dart.stopwatch_elapsed_microseconds" => {
            crate::emitter::string_adapter::emit_dart_stopwatch_elapsed_microseconds(
                chunks, current, line,
            )
        }
        "dart.future_call0" => {
            crate::emitter::string_adapter::emit_dart_future_call0(chunks, current, line)
        }
        "dart.future_delayed" => {
            crate::emitter::string_adapter::emit_dart_future_delayed(chunks, current, line)
        }
        "dart.contains" => {
            crate::emitter::string_adapter::emit_dart_contains(chunks, current, line)
        }
        "dart.iter_element_at" => {
            crate::emitter::string_adapter::emit_dart_iter_element_at(chunks, current, line)
        }
        "dart.list_index_where" => {
            crate::emitter::string_adapter::emit_dart_list_where_search(chunks, current, 0, line)
        }
        "dart.list_last_index_where" => {
            crate::emitter::string_adapter::emit_dart_list_where_search(chunks, current, 1, line)
        }
        "dart.list_first_where" => {
            crate::emitter::string_adapter::emit_dart_list_where_search(chunks, current, 2, line)
        }
        "dart.list_last_where" => {
            crate::emitter::string_adapter::emit_dart_list_where_search(chunks, current, 3, line)
        }
        "dart.map_general" => {
            crate::emitter::string_adapter::emit_dart_map_general(chunks, current, line)
        }
        "dart.iter_async_map" => {
            crate::emitter::string_adapter::emit_dart_iter_async_map(chunks, current, line)
        }
        "dart.iter_where" => {
            crate::emitter::string_adapter::emit_dart_iter_where(chunks, current, line)
        }
        "dart.iter_expand" => {
            crate::emitter::string_adapter::emit_dart_iter_expand(chunks, current, line)
        }
        "dart.iter_expand_precurrent" => {
            crate::emitter::string_adapter::emit_dart_iter_expand_precurrent(chunks, current, line)
        }
        "dart.iter_followed_by" => {
            crate::emitter::string_adapter::emit_dart_iter_followed_by(chunks, current, line)
        }
        "dart.iter_take" => {
            crate::emitter::string_adapter::emit_dart_iter_take(chunks, current, line)
        }
        "dart.iter_skip" => {
            crate::emitter::string_adapter::emit_dart_iter_skip(chunks, current, line)
        }
        "dart.iter_take_while" => {
            crate::emitter::string_adapter::emit_dart_iter_take_while(chunks, current, line)
        }
        "dart.iter_skip_while" => {
            crate::emitter::string_adapter::emit_dart_iter_skip_while(chunks, current, line)
        }
        "dart.iter_distinct" => {
            crate::emitter::string_adapter::emit_dart_iter_distinct(chunks, current, argc, line)
        }
        "dart.iter_any" => {
            crate::emitter::string_adapter::emit_dart_iter_any(chunks, current, line)
        }
        "dart.iter_every" => {
            crate::emitter::string_adapter::emit_dart_iter_every(chunks, current, line)
        }
        "dart.iter_reduce" => {
            crate::emitter::string_adapter::emit_dart_iter_reduce(chunks, current, argc, line)
        }
        "dart.iter_for_each" => {
            crate::emitter::string_adapter::emit_dart_for_each_general(chunks, current, line)
        }
        "dart.print" => {
            crate::emitter::string_adapter::emit_dart_print(chunks, current, argc, line)
        }
        "dart.to_string" => {
            crate::emitter::string_adapter::emit_dart_to_string(chunks, current, line)
        }
        "dart.double_to_string" => {
            crate::emitter::string_adapter::emit_dart_double_to_string(chunks, current, line)
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
