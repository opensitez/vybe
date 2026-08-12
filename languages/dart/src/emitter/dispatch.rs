//! Auto-extracted `dart.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_compiler::primitives::url::UrlField;
use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        // Structural equality for composite built-ins — Dart records/tuples
        // compare by VALUE. Reached as `[builtin_slots.array] eq`; used to be
        // the `LanguageHooks::value_eq` callback, keyed by language name.
        "dart.value_eq" => {
            vybe_compiler::primitives::tuples::emit_tuple_value_eq(&mut chunks[current], line)
        }
        // dart:io filesystem — see emitter/io_adapter.rs
        "dart.io_read_as_string_sync" => {
            crate::emitter::io_adapter::emit_read_as_string_sync(chunks, current, argc, line)
        }
        "dart.io_read_as_latin1_string_sync" => {
            crate::emitter::io_adapter::emit_read_as_latin1_string_sync(chunks, current, argc, line)
        }
        "dart.io_read_as_bytes_sync" => {
            crate::emitter::io_adapter::emit_read_as_bytes_sync(chunks, current, argc, line)
        }
        "dart.io_read_as_lines_sync" => {
            crate::emitter::io_adapter::emit_read_as_lines_sync(chunks, current, argc, line)
        }
        "dart.io_write_as_string_sync" => {
            crate::emitter::io_adapter::emit_write_as_string_sync(chunks, current, argc, line)
        }
        "dart.io_write_as_bytes_sync" => {
            crate::emitter::io_adapter::emit_write_as_bytes_sync(chunks, current, argc, line)
        }
        "dart.io_append_as_string_sync" => {
            crate::emitter::io_adapter::emit_append_as_string_sync(chunks, current, argc, line)
        }
        "dart.io_append_as_bytes_sync" => {
            crate::emitter::io_adapter::emit_append_as_bytes_sync(chunks, current, argc, line)
        }
        "dart.io_exists_sync" => {
            crate::emitter::io_adapter::emit_exists_sync(chunks, current, argc, line)
        }
        "dart.io_delete_sync" => {
            crate::emitter::io_adapter::emit_delete_sync(chunks, current, argc, line)
        }
        "dart.io_length_sync" => {
            crate::emitter::io_adapter::emit_length_sync(chunks, current, argc, line)
        }
        "dart.io_create_sync" => {
            crate::emitter::io_adapter::emit_create_sync(chunks, current, argc, line)
        }
        "dart.io_rename_sync" => {
            crate::emitter::io_adapter::emit_rename_sync(chunks, current, argc, line)
        }
        "dart.io_copy_sync" => {
            crate::emitter::io_adapter::emit_copy_sync(chunks, current, argc, line)
        }
        "dart.io_list_sync" => {
            crate::emitter::io_adapter::emit_list_sync(chunks, current, argc, line)
        }
        "dart.io_stat_sync" => {
            crate::emitter::io_adapter::emit_stat_sync(chunks, current, argc, line)
        }
        "dart.io_stat_path" => {
            crate::emitter::io_adapter::emit_stat_path(chunks, current, argc, line)
        }
        "dart.io_last_modified_sync" => {
            crate::emitter::io_adapter::emit_last_modified_sync(chunks, current, argc, line)
        }
        "dart.io_set_last_modified_sync" => {
            crate::emitter::io_adapter::emit_set_last_modified_sync(chunks, current, argc, line)
        }
        "dart.io_set_last_accessed_sync" => {
            crate::emitter::io_adapter::emit_set_last_accessed_sync(chunks, current, argc, line)
        }
        "dart.io_resolve_symbolic_links_sync" => {
            crate::emitter::io_adapter::emit_resolve_symbolic_links_sync(chunks, current, argc, line)
        }
        "dart.io_target_sync" => {
            crate::emitter::io_adapter::emit_target_sync(chunks, current, argc, line)
        }
        "dart.io_update_sync" => {
            crate::emitter::io_adapter::emit_update_sync(chunks, current, argc, line)
        }
        "dart.io_create_temp_sync" => {
            crate::emitter::io_adapter::emit_create_temp_sync(chunks, current, argc, line)
        }
        "dart.io_watch" => {
            crate::emitter::io_adapter::emit_watch(chunks, current, argc, line)
        }
        "dart.io_absolute_handle" => {
            crate::emitter::io_adapter::emit_absolute_handle(chunks, current, argc, line)
        }
        "dart.io_parent_handle" => {
            crate::emitter::io_adapter::emit_parent_handle(chunks, current, argc, line)
        }
        "dart.io_uri_string" => {
            crate::emitter::io_adapter::emit_uri_string(chunks, current, argc, line)
        }
        "dart.io_is_absolute" => {
            crate::emitter::io_adapter::emit_is_absolute(chunks, current, argc, line)
        }
        "dart.io_handle_is_absolute" => {
            crate::emitter::io_adapter::emit_handle_is_absolute(chunks, current, argc, line)
        }
        "dart.io_type_sync" => {
            crate::emitter::io_adapter::emit_type_sync(chunks, current, argc, line)
        }
        "dart.io_identical_sync" => {
            crate::emitter::io_adapter::emit_identical_sync(chunks, current, argc, line)
        }
        "dart.io_set_current_dir" => {
            crate::emitter::io_adapter::emit_set_current_dir(chunks, current, argc, line)
        }
        "dart.io_platform_environment" => {
            crate::emitter::io_adapter::emit_platform_environment(chunks, current, line)
        }
        "dart.io_process_run_sync" => {
            crate::emitter::io_adapter::emit_process_run_sync(chunks, current, argc, line)
        }
        "dart.io_process_start" => {
            crate::emitter::io_adapter::emit_process_start(chunks, current, argc, line)
        }
        "dart.io_process_kill" => {
            crate::emitter::io_adapter::emit_process_kill(chunks, current, argc, line)
        }
        "dart.io_process_stdin_writeln" => {
            crate::emitter::io_adapter::emit_process_stdin_writeln(chunks, current, argc, line)
        }
        "dart.io_process_stdin_add" => {
            crate::emitter::io_adapter::emit_process_stdin_add(chunks, current, argc, line)
        }
        "dart.io_process_stdin_write_char_code" => {
            crate::emitter::io_adapter::emit_process_stdin_write_char_code(chunks, current, argc, line)
        }
        "dart.io_process_stdin_flush" => {
            crate::emitter::io_adapter::emit_process_stdin_flush(chunks, current, argc, line)
        }
        "dart.io_process_stdin_close" => {
            crate::emitter::io_adapter::emit_process_stdin_close(chunks, current, argc, line)
        }
        "dart.io_process_stdin_add_error" => {
            crate::emitter::io_adapter::emit_process_stdin_add_error(chunks, current, argc, line)
        }
        "dart.utf8_encode" => {
            crate::emitter::io_adapter::emit_utf8_encode(chunks, current, argc, line)
        }
        "dart.latin1_encode" => {
            crate::emitter::io_adapter::emit_latin1_encode(chunks, current, argc, line)
        }
        "dart.utf8_decode" => {
            crate::emitter::io_adapter::emit_utf8_decode(chunks, current, argc, line)
        }
        "dart.latin1_decode" => {
            crate::emitter::io_adapter::emit_latin1_decode(chunks, current, argc, line)
        }
        "dart.io_open_sync" => {
            crate::emitter::io_adapter::emit_open_sync(chunks, current, argc, line)
        }
        "dart.io_raf_close_sync" => {
            crate::emitter::io_adapter::emit_raf_close_sync(chunks, current, argc, line)
        }
        "dart.io_raf_flush_sync" => {
            crate::emitter::io_adapter::emit_raf_flush_sync(chunks, current, argc, line)
        }
        "dart.io_raf_lock_sync" => {
            crate::emitter::io_adapter::emit_raf_lock_sync(chunks, current, argc, line)
        }
        "dart.io_raf_unlock_sync" => {
            crate::emitter::io_adapter::emit_raf_unlock_sync(chunks, current, argc, line)
        }
        "dart.io_raf_length_sync" => {
            crate::emitter::io_adapter::emit_raf_length_sync(chunks, current, argc, line)
        }
        "dart.io_raf_truncate_sync" => {
            crate::emitter::io_adapter::emit_raf_truncate_sync(chunks, current, argc, line)
        }
        "dart.io_raf_write_string_sync" => {
            crate::emitter::io_adapter::emit_raf_write_string_sync(chunks, current, argc, line)
        }
        "dart.io_raf_write_byte_sync" => {
            crate::emitter::io_adapter::emit_raf_write_byte_sync(chunks, current, argc, line)
        }
        "dart.io_raf_write_from_sync" => {
            crate::emitter::io_adapter::emit_raf_write_from_sync(chunks, current, argc, line)
        }
        "dart.io_raf_read_byte_sync" => {
            crate::emitter::io_adapter::emit_raf_read_byte_sync(chunks, current, argc, line)
        }
        "dart.io_raf_read_sync" => {
            crate::emitter::io_adapter::emit_raf_read_sync(chunks, current, argc, line)
        }
        "dart.io_raf_read_into_sync" => {
            crate::emitter::io_adapter::emit_raf_read_into_sync(chunks, current, argc, line)
        }
        "dart.io_raf_position_sync" => {
            crate::emitter::io_adapter::emit_raf_position_sync(chunks, current, argc, line)
        }
        "dart.io_raf_set_position_sync" => {
            crate::emitter::io_adapter::emit_raf_set_position_sync(chunks, current, argc, line)
        }
        "dart.is_empty" => {
            crate::emitter::string_adapter::emit_dart_is_empty(chunks, current, line)
        }
        "dart.is_not_empty" => {
            crate::emitter::string_adapter::emit_dart_is_not_empty(chunks, current, line)
        }
        "dart.is_even" => crate::emitter::string_adapter::emit_dart_is_even(chunks, current, line),
        "dart.is_odd" => crate::emitter::string_adapter::emit_dart_is_odd(chunks, current, line),
        "dart.sb_new" => {
            crate::emitter::string_adapter::emit_dart_sb_new(chunks, current, argc, line)
        }
        "dart.sb_to_string" => {
            crate::emitter::string_adapter::emit_dart_sb_to_string(chunks, current, line)
        }
        "dart.sb_length" => {
            crate::emitter::string_adapter::emit_dart_sb_length(chunks, current, line)
        }
        "dart.sb_is_empty" => {
            crate::emitter::string_adapter::emit_dart_sb_is_empty(chunks, current, line)
        }
        "dart.sb_is_not_empty" => {
            crate::emitter::string_adapter::emit_dart_sb_is_not_empty(chunks, current, line)
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
        // `dart.exception` / `.format_exception` / `.range_error` /
        // `.state_error` / `.argument_error` / `.unimplemented_error` are GONE —
        // those six types are `dart:core` classes now
        // (`core_classes/exceptions.rs`), constructed through `ExprKind::New`.
        "dart.stack_trace" => {
            crate::emitter::string_adapter::emit_dart_stack_trace(chunks, current, line)
        }
        "dart.base64_encode" => {
            crate::emitter::string_adapter::emit_dart_base64_encode(chunks, current, false, line)
        }
        "dart.base64_decode" => {
            crate::emitter::string_adapter::emit_dart_base64_decode(chunks, current, false, line)
        }
        "dart.base64_normalize" => {
            crate::emitter::string_adapter::emit_dart_base64_normalize(chunks, current, false, line)
        }
        "dart.base64url_encode" => {
            crate::emitter::string_adapter::emit_dart_base64_encode(chunks, current, true, line)
        }
        "dart.base64url_decode" => {
            crate::emitter::string_adapter::emit_dart_base64_decode(chunks, current, true, line)
        }
        "dart.base64url_normalize" => {
            crate::emitter::string_adapter::emit_dart_base64_normalize(chunks, current, true, line)
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
        // The `dart.datetime_*` family is GONE — `DateTime` is a class
        // (`core_classes/datetime.rs`) and every member is spelled in Dart.
        // What is left of the domain is two CONVENTIONS the AST names.
        "dart.date_month" => crate::emitter::core_adapter::emit_date_month(chunks, current, line),
        "dart.date_weekday" => {
            crate::emitter::core_adapter::emit_date_weekday(chunks, current, line)
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
        "dart.compare_to" => crate::emitter::core_adapter::emit_compare_to(chunks, current, line),
        // The `dart.uri_*` family is GONE. `Uri` is a class
        // (`core_classes/uri.rs`) and every member below is spelled in Dart in
        // its body, so `normalizePath`/`replace`/`resolve` are real code rather
        // than the empty `{}` emitters two of them used to be. What remains is
        // the parse and the component reads, which `primitives::url` owns.
        "dart.url_parse" => crate::emitter::core_adapter::emit_url_parse(chunks, current, line),
        "dart.url_scheme" => crate::emitter::core_adapter::emit_url_component(
            chunks,
            current,
            UrlField::Scheme,
            line,
        ),
        "dart.url_host" => {
            crate::emitter::core_adapter::emit_url_component(chunks, current, UrlField::Host, line)
        }
        "dart.url_port" => {
            crate::emitter::core_adapter::emit_url_component(chunks, current, UrlField::Port, line)
        }
        "dart.url_path" => {
            crate::emitter::core_adapter::emit_url_component(chunks, current, UrlField::Path, line)
        }
        "dart.url_query" => {
            crate::emitter::core_adapter::emit_url_component(chunks, current, UrlField::Query, line)
        }
        "dart.url_fragment" => crate::emitter::core_adapter::emit_url_component(
            chunks,
            current,
            UrlField::Fragment,
            line,
        ),
        "dart.url_user" => {
            crate::emitter::core_adapter::emit_url_component(chunks, current, UrlField::User, line)
        }
        "dart.url_pass" => {
            crate::emitter::core_adapter::emit_url_component(chunks, current, UrlField::Pass, line)
        }
        "dart.url_decode" => crate::emitter::core_adapter::emit_url_decode(chunks, current, line),
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
            vybe_compiler::primitives::sorted_collection::emit_sorted_add(chunks, current, line)
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
        "dart.stream_as_broadcast" => {
            crate::emitter::string_adapter::emit_dart_stream_as_broadcast(chunks, current, line)
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
        "dart.bigint_compare_to" => {
            crate::emitter::string_adapter::emit_dart_bigint_compare_to(chunks, current, line)
        }
        "dart.bigint_idiv" => {
            crate::emitter::string_adapter::emit_dart_bigint_idiv(chunks, current, line)
        }
        "dart.bigint_mod" => {
            crate::emitter::string_adapter::emit_dart_bigint_mod(chunks, current, line)
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
        "dart.nullable_double_to_string" => {
            crate::emitter::string_adapter::emit_dart_nullable_double_to_string(
                chunks, current, line,
            )
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
