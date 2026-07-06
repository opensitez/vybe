//! Auto-extracted `dotnet.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_emitter::instructions::core_wasm;
use vybe_bytecode::Chunk;

use vybe_bytecode::opcode::Op;

/// VB `Choose(idx, v1, v2, ..., vN)` — variadic 1-indexed selector.
/// Packs the trailing values into an array, then `ARRAY_GET array[idx-1]`.
fn emit_choose(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc < 2 {
        chunk.emit_op(Op::NULL, line);
        return;
    }
    let n = (argc as u16) - 1;
    let arr_slot = chunk.alloc_scratch(2);
    let idx_slot = arr_slot + 1;

    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, n, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if crate::emitter::core::runtime_adapter::emit_helper(name, chunks, current, argc, line)
    {
        return true;
    }
    match name {
        "dotnet.dns_get_host_addresses" => {
            crate::emitter::core::sockets_adapter::emit_dns_get_host_addresses(
                chunks, current, line,
            )
        }
        "dotnet.dns_get_host_entry" => {
            crate::emitter::core::sockets_adapter::emit_dns_get_host_entry(
                chunks, current, line,
            )
        }
        "dotnet.dns_get_host_name" => {
            crate::emitter::core::sockets_adapter::emit_dns_get_host_name(
                chunks, current, line,
            )
        }
        "dotnet.tcp_client_new" => {
            crate::emitter::core::sockets_adapter::emit_tcp_client_new(
                chunks, current, line,
            )
        }
        "dotnet.tcp_client_get_stream" => {
            crate::emitter::core::sockets_adapter::emit_tcp_client_get_stream(
                chunks, current, line,
            )
        }
        "dotnet.tcp_client_close" => {
            crate::emitter::core::sockets_adapter::emit_tcp_client_close(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_new" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_new(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_start" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_start(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_stop" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_stop(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_accept" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_accept(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_pending" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_pending(
                chunks, current, line,
            )
        }
        "dotnet.udp_client_new" => {
            crate::emitter::core::sockets_adapter::emit_udp_client_new(
                chunks, current, line,
            )
        }
        "dotnet.udp_send" => {
            crate::emitter::core::sockets_adapter::emit_udp_send(chunks, current, line)
        }
        "dotnet.udp_receive" => {
            crate::emitter::core::sockets_adapter::emit_udp_receive(chunks, current, line)
        }
        "dotnet.udp_close" => {
            crate::emitter::core::sockets_adapter::emit_udp_close(chunks, current, line)
        }

        // ── .NET StringBuilder adapter ──────────────────────────────
        // No direct ECMA mirror; the wrapper materializes a plain
        // Object with a `__buffer` string and mutates via DYN_ADD +
        // STRUCT_SET. Multi-arity ctor uses the threaded `argc` to
        // pick between empty / initial-keyed shapes.
        "dotnet.string_builder_new" => {
            crate::emitter::core::stringbuilder_adapter::emit_string_builder_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.sb_append" => crate::emitter::core::stringbuilder_adapter::emit_sb_append(
            chunks, current, line,
        ),
        "dotnet.sb_append_line" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_append_line(
                chunks, current, argc, line,
            )
        }
        "dotnet.sb_append_format" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_append_format(
                chunks, current, argc, line,
            )
        }
        "dotnet.sb_to_string" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_to_string(
                chunks, current, argc, line,
            )
        }
        "dotnet.sb_clear" => crate::emitter::core::stringbuilder_adapter::emit_sb_clear(
            chunks, current, line,
        ),
        "dotnet.sb_length" => crate::emitter::core::stringbuilder_adapter::emit_sb_length(
            chunks, current, line,
        ),
        "dotnet.sb_insert" => crate::emitter::core::stringbuilder_adapter::emit_sb_insert(
            chunks, current, line,
        ),
        "dotnet.sb_remove" => crate::emitter::core::stringbuilder_adapter::emit_sb_remove(
            chunks, current, line,
        ),
        "dotnet.sb_replace" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_replace(
                chunks, current, line,
            )
        }
        "dotnet.sb_index_get" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_index_get(
                chunks, current, line,
            )
        }
        "dotnet.sb_index_set" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_index_set(
                chunks, current, line,
            )
        }

        // ── .NET Random adapter ─────────────────────────────────────
        "dotnet.random_new" => crate::emitter::core::random_adapter::emit_random_new(
            chunks, current, argc, line,
        ),
        "dotnet.random_next" => crate::emitter::core::random_adapter::emit_random_next(
            chunks, current, argc, line,
        ),
        "dotnet.random_next_double" => {
            crate::emitter::core::random_adapter::emit_random_next_double(
                chunks, current, line,
            )
        }

        // ── .NET Regex adapter ──────────────────────────────────────
        "dotnet.regex_new" => {
            crate::emitter::core::regex_adapter::emit_regex_new(chunks, current, argc, line)
        }
        "dotnet.regex_is_match" => {
            crate::emitter::core::regex_adapter::emit_regex_is_match(chunks, current, line)
        }
        "dotnet.regex_replace" => {
            crate::emitter::core::regex_adapter::emit_regex_replace(chunks, current, line)
        }
        "dotnet.regex_split" => {
            crate::emitter::core::regex_adapter::emit_regex_split(chunks, current, line)
        }
        "dotnet.regex_match" => {
            crate::emitter::core::regex_adapter::emit_regex_match(chunks, current, line)
        }
        "dotnet.regex_matches" => {
            crate::emitter::core::regex_adapter::emit_regex_matches(chunks, current, line)
        }

        // ── .NET Stopwatch adapter ──────────────────────────────────
        "dotnet.stopwatch_new" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_new(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_start" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_start(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_stop" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_stop(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_reset" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_reset(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_start_new" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_start_new(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_restart" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_restart(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_elapsed_ms" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_elapsed_ms(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_is_running" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_is_running(
                chunks, current, line,
            )
        }

        // ── .NET Process / ProcessStartInfo adapter ─────────────────
        // Lowers to `node:child_process.spawnSync` + plain Object
        // structs for the .NET-shape records. Multi-arity ctors use
        // the threaded `argc`.
        "dotnet.process_start_info_new" => {
            crate::emitter::core::process_adapter::emit_process_start_info_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.process_new" => crate::emitter::core::process_adapter::emit_process_new(
            chunks, current, argc, line,
        ),
        "dotnet.process_start" => {
            crate::emitter::core::process_adapter::emit_process_start(chunks, current, line)
        }
        "dotnet.process_get_current" => {
            crate::emitter::core::process_adapter::emit_process_get_current(
                chunks, current, line,
            )
        }
        "dotnet.process_wait_for_exit" => {
            crate::emitter::core::process_adapter::emit_process_wait_for_exit(
                chunks, current, line,
            )
        }

        // ── .NET System.Array static-method adapter ─────────────────
        // `Clear` / `Copy` / `Resize` / `Sort` lower to bundled stdlib
        // chunks (`__vybe_*` globals) composing `ecma:array.*`
        // primitives. No `vybe:types/array*` host fns.
        "dotnet.array_clear" => {
            crate::emitter::core::array_adapter::emit_array_clear(chunks, current, line)
        }
        "dotnet.array_copy" => {
            crate::emitter::core::array_adapter::emit_array_copy(chunks, current, line)
        }
        "dotnet.array_resize" => {
            crate::emitter::core::array_adapter::emit_array_resize(chunks, current, line)
        }
        "dotnet.array_sort" => {
            crate::emitter::core::array_adapter::emit_array_sort(chunks, current, line)
        }
        "dotnet.hashset_add" => {
            crate::emitter::core::collections_adapter::emit_hashset_add(
                chunks, current, line,
            )
        }
        "dotnet.set_new_ignore_comparer" => {
            crate::emitter::core::collections_adapter::emit_set_new_ignore_comparer(
                chunks, current, line,
            )
        }
        "dotnet.hashset_union_with" => {
            crate::emitter::core::collections_adapter::emit_hashset_union_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_intersect_with" => {
            crate::emitter::core::collections_adapter::emit_hashset_intersect_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_except_with" => {
            crate::emitter::core::collections_adapter::emit_hashset_except_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_symmetric_except_with" => {
            crate::emitter::core::collections_adapter::emit_hashset_symmetric_except_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_is_subset_of" => {
            crate::emitter::core::collections_adapter::emit_hashset_is_subset_of(
                chunks, current, line,
            )
        }
        "dotnet.hashset_is_superset_of" => {
            crate::emitter::core::collections_adapter::emit_hashset_is_superset_of(
                chunks, current, line,
            )
        }
        "dotnet.hashset_overlaps" => {
            crate::emitter::core::collections_adapter::emit_hashset_overlaps(chunks, current, line)
        }
        "dotnet.task_wait" => {
            crate::emitter::core::thread_adapter::emit_task_wait(chunks, current, line)
        }
        "dotnet.hashset_set_equals" => {
            crate::emitter::core::collections_adapter::emit_hashset_set_equals(
                chunks, current, line,
            )
        }
        "dotnet.hashset_is_proper_subset_of" => {
            crate::emitter::core::collections_adapter::emit_hashset_is_proper_subset_of(
                chunks, current, line,
            )
        }
        "dotnet.hashset_is_proper_superset_of" => {
            crate::emitter::core::collections_adapter::emit_hashset_is_proper_superset_of(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_add_first" => {
            crate::emitter::core::collections_adapter::emit_linked_list_add_first(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_add_last" => {
            crate::emitter::core::collections_adapter::emit_linked_list_add_last(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_find" => {
            crate::emitter::core::collections_adapter::emit_linked_list_find(
                chunks, current, line,
            )
        }
        "dotnet.sorted_dictionary_entries" => {
            crate::emitter::core::collections_adapter::emit_sorted_dictionary_entries(
                chunks, current, line,
            )
        }

        // ── .NET TimeSpan factory adapters ──────────────────────────
        // `TimeSpan.From*(n)` factories build a duration record by
        // multiplying `n` with the unit-to-ms factor. Pure inline
        // bytecode; no host fns.
        "dotnet.timespan_from_days" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_days(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_hours" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_hours(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_minutes" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_minutes(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_seconds" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_seconds(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_milliseconds" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_milliseconds(
                chunks, current, line,
            )
        }
        "dotnet.timespan_zero" => {
            crate::emitter::core::timespan_adapter::emit_timespan_zero(
                chunks, current, line,
            )
        }
        "dotnet.timespan_new" => crate::emitter::core::timespan_adapter::emit_timespan_new(
            chunks, current, argc, line,
        ),
        "dotnet.timespan_compare" => {
            crate::emitter::core::timespan_adapter::emit_timespan_compare(
                chunks, current, line,
            )
        }
        "dotnet.timespan_negate" => {
            crate::emitter::core::timespan_adapter::emit_timespan_negate(
                chunks, current, line,
            )
        }
        "dotnet.timespan_duration" => {
            crate::emitter::core::timespan_adapter::emit_timespan_duration(
                chunks, current, line,
            )
        }
        "dotnet.timespan_add" => {
            crate::emitter::core::timespan_adapter::emit_timespan_add(chunks, current, line)
        }
        "dotnet.timespan_sub" => {
            crate::emitter::core::timespan_adapter::emit_timespan_sub(chunks, current, line)
        }

        // ── .NET Guid adapters ──────────────────────────────────────
        // `Guid` is stored as a .NET-shaped object carrying the
        // canonical lowercase text representation in `__value`.
        "dotnet.guid_empty" => {
            crate::emitter::core::guid_adapter::emit_guid_empty(chunks, current, line)
        }
        "dotnet.guid_new_guid" => {
            crate::emitter::core::guid_adapter::emit_guid_new_guid(chunks, current, line)
        }
        "dotnet.guid_parse" => {
            crate::emitter::core::guid_adapter::emit_guid_parse(chunks, current, line)
        }
        "dotnet.guid_new" => {
            crate::emitter::core::guid_adapter::emit_guid_new(chunks, current, argc, line)
        }
        "dotnet.guid_to_string" => {
            crate::emitter::core::guid_adapter::emit_guid_to_string(chunks, current, line)
        }
        "dotnet.guid_try_parse" => crate::emitter::core::guid_adapter::emit_guid_try_parse(
            chunks, current, argc, line,
        ),

        "dotnet.version_new" => crate::emitter::core::version_adapter::emit_version_new(
            chunks, current, argc, line,
        ),
        "dotnet.version_parse" => {
            crate::emitter::core::version_adapter::emit_version_parse(chunks, current, line)
        }
        "dotnet.version_to_string" => {
            crate::emitter::core::version_adapter::emit_version_to_string(
                chunks, current, line,
            )
        }
        "dotnet.version_compare" => {
            crate::emitter::core::version_adapter::emit_version_compare(
                chunks, current, line,
            )
        }
        "dotnet.version_equals" => {
            crate::emitter::core::version_adapter::emit_version_equals(
                chunks, current, line,
            )
        }
        "dotnet.version_lt" => {
            crate::emitter::core::version_adapter::emit_version_lt(chunks, current, line)
        }
        "dotnet.version_gt" => {
            crate::emitter::core::version_adapter::emit_version_gt(chunks, current, line)
        }
        "dotnet.version_eq" => {
            crate::emitter::core::version_adapter::emit_version_eq(chunks, current, line)
        }
        "dotnet.version_ne" => {
            crate::emitter::core::version_adapter::emit_version_ne(chunks, current, line)
        }

        // ── .NET DateTime static adapters ───────────────────────────
        // `Now` / `UtcNow` / `Today` lower to `ecma:date.now` (which
        // reads `wasi:clocks/wall-clock.now`); `Parse` lowers to
        // `ecma:date.parse`. Each wraps the resulting ms timestamp
        // in a `{__type:"DateTime", __time:ms}` object so the .NET
        // surface looks .NET-shaped.
        "dotnet.datetime_now" => {
            crate::emitter::core::datetime_adapter::emit_datetime_now(chunks, current, line)
        }
        "dotnet.datetime_parse" => {
            crate::emitter::core::datetime_adapter::emit_datetime_parse(
                chunks, current, line,
            )
        }
        "dotnet.datetime_today" => {
            crate::emitter::core::datetime_adapter::emit_datetime_today(
                chunks, current, line,
            )
        }
        "dotnet.datetime_new" => crate::emitter::core::datetime_adapter::emit_datetime_new(
            chunks, current, argc, line,
        ),
        "dotnet.datetime_add_days" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_days(
                chunks, current, line,
            )
        }
        "dotnet.datetime_add_hours" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_hours(
                chunks, current, line,
            )
        }
        "dotnet.datetime_add_months" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_months(
                chunks, current, line,
            )
        }
        "dotnet.datetime_days_in_month" => {
            crate::emitter::core::datetime_adapter::emit_datetime_days_in_month(
                chunks, current, line,
            )
        }
        "dotnet.datetime_is_leap_year" => {
            crate::emitter::core::datetime_adapter::emit_datetime_is_leap_year(
                chunks, current, line,
            )
        }
        "dotnet.datetime_compare" => {
            crate::emitter::core::datetime_adapter::emit_datetime_compare(
                chunks, current, line,
            )
        }
        "dotnet.datetime_to_short_date_string" => {
            crate::emitter::core::datetime_adapter::emit_datetime_to_short_date_string(
                chunks, current, line,
            )
        }
        "dotnet.datetime_add_timespan" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_timespan(
                chunks, current, line,
            )
        }
        "dotnet.datetime_subtract_datetime" => {
            crate::emitter::core::datetime_adapter::emit_datetime_subtract_datetime(
                chunks, current, line,
            )
        }

        // ── PHP DateTime / DateTimeImmutable / DateInterval adapters ──
        // Bytecode-only — composes existing `ecma:date.*` host fns into
        // the PHP-shaped surface. See `emitter/php/datetime_adapter.rs`.
        "dotnet.string_format" => {
            crate::emitter::core::string_format_adapter::emit_string_format(
                chunks, current, argc, line,
            )
        }

        // ── VB / VBA `Format(value, picture)` — picture-string render ──
        "dotnet.format_picture" => {
            crate::emitter::core::format_picture_adapter::emit_format_picture(
                chunks, current, argc, line,
            )
        }

        // ── .NET StreamReader / StreamWriter adapters — text I/O ────
        // Load-whole-file model: `new StreamReader(path)` materializes a
        // string buffer via `node:fs.readFileSync`, `new StreamWriter`
        // accumulates into `__buf` and flushes via `writeFileSync`.
        // Bytecode-only — no `dotnet:io` host fns.
        "dotnet.stream_reader_new" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.stream_reader_read_line" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_read_line(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_read_to_end" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_read_to_end(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_at_end" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_at_end(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_close" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_close(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_new" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.stream_writer_write" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_write(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_write_line" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_write_line(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_flush" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_flush(
                chunks, current, line,
            )
        }
        "dotnet.stream_close" => {
            crate::emitter::core::stream_io_adapter::emit_stream_close(
                chunks, current, line,
            )
        }
        "dotnet.file_read_all_lines" => {
            crate::emitter::core::filesystem_adapter::emit_file_read_all_lines(
                chunks, current, line,
            )
        }
        "dotnet.directory_get_files" => {
            crate::emitter::core::filesystem_adapter::emit_directory_get_files(
                chunks, current, line,
            )
        }
        "dotnet.directory_get_directories" => {
            crate::emitter::core::filesystem_adapter::emit_directory_get_directories(
                chunks, current, line,
            )
        }
        "dotnet.console_writeline" => {
            crate::emitter::core::console_adapter::emit_console_writeline(
                chunks, current, line,
            )
        }
        "dotnet.console_readline" => {
            crate::emitter::core::console_adapter::emit_console_readline(
                chunks, current, line,
            )
        }
        "dotnet.console_error" => {
            crate::emitter::core::console_adapter::emit_console_error(chunks, current, line)
        }
        "dotnet.environment_username" => {
            crate::emitter::core::environment_adapter::emit_environment_username(
                chunks, current, line,
            )
        }
        "dotnet.environment_processor_count" => {
            crate::emitter::core::environment_adapter::emit_environment_processor_count(
                chunks, current, line,
            )
        }
        "dotnet.environment_tick_count" => {
            crate::emitter::core::environment_adapter::emit_environment_tick_count(
                chunks, current, line,
            )
        }
        "dotnet.environment_get" => {
            crate::emitter::core::environment_adapter::emit_environment_get(
                chunks, current, line,
            )
        }
        "dotnet.environment_set" => {
            crate::emitter::core::environment_adapter::emit_environment_set(
                chunks, current, line,
            )
        }

        // ── OleDb adapter — System.Data.OleDb constructor wrappers ─────────────
        "dotnet.oledb_connection_new" => {
            crate::emitter::core::oledb_adapter::emit_oledb_connection_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.oledb_command_new" => {
            crate::emitter::core::oledb_adapter::emit_oledb_command_new(
                chunks, current, argc, line,
            )
        }

        // ── ADODB adapter — ADODB.Connection / Command / Recordset ──────────────
        "dotnet.adodb_connection_new" => {
            crate::emitter::core::adodb_adapter::emit_adodb_connection_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_connection_execute" => {
            crate::emitter::core::adodb_adapter::emit_adodb_connection_execute(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_begin_trans" => {
            crate::emitter::core::adodb_adapter::emit_adodb_conn_begin_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_commit_trans" => {
            crate::emitter::core::adodb_adapter::emit_adodb_conn_commit_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_rollback_trans" => {
            crate::emitter::core::adodb_adapter::emit_adodb_conn_rollback_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_command_new" => {
            crate::emitter::core::adodb_adapter::emit_adodb_command_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_command_execute" => {
            crate::emitter::core::adodb_adapter::emit_adodb_command_execute(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_command_create_parameter" => {
            crate::emitter::core::adodb_adapter::emit_adodb_command_create_parameter(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_new" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_open" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_open(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_move_next" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_move_next(
                chunks, current, line,
            )
        }
        "dotnet.adodb_recordset_move_first" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_move_first(
                chunks, current, line,
            )
        }
        "dotnet.adodb_recordset_fields" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_fields(
                chunks, current, line,
            )
        }
        "dotnet.adodb_recordset_close" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_close(
                chunks, current, line,
            )
        }

        // ── LINQ surface — composed bytecode shared by every .NET-shape language ──
        "dotnet.linq_first" => {
            crate::emitter::core::linq_adapter::emit_linq_first(chunks, current, line)
        }
        "dotnet.linq_last" => {
            crate::emitter::core::linq_adapter::emit_linq_last(chunks, current, line)
        }
        "dotnet.linq_skip" => {
            crate::emitter::core::linq_adapter::emit_linq_skip(chunks, current, line)
        }
        "dotnet.linq_take" => {
            crate::emitter::core::linq_adapter::emit_linq_take(chunks, current, line)
        }
        "dotnet.linq_identity" => {
            crate::emitter::core::linq_adapter::emit_linq_identity(chunks, current, line)
        }
        "dotnet.linq_average" => {
            crate::emitter::core::linq_adapter::emit_linq_average(chunks, current, line)
        }
        "dotnet.linq_first_or_default" => {
            crate::emitter::core::linq_adapter::emit_linq_first_or_default(
                chunks, current, line,
            )
        }
        "dotnet.linq_distinct" => {
            crate::emitter::core::linq_adapter::emit_linq_distinct(chunks, current, line)
        }
        "dotnet.linq_distinct_by" => {
            crate::emitter::core::linq_adapter::emit_linq_distinct_by(chunks, current, line)
        }
        "dotnet.linq_order_by" => {
            crate::emitter::core::linq_adapter::emit_linq_order_by(chunks, current, line)
        }
        "dotnet.linq_sequence_equal" => {
            crate::emitter::core::linq_adapter::emit_linq_sequence_equal(
                chunks, current, line,
            )
        }
        "dotnet.linq_count_pred" => {
            crate::emitter::core::linq_adapter::emit_linq_count_pred(chunks, current, line)
        }
        "dotnet.linq_aggregate" => {
            crate::emitter::core::linq_adapter::emit_linq_aggregate(chunks, current, line)
        }
        "dotnet.linq_order_by_descending" => {
            crate::emitter::core::linq_adapter::emit_linq_order_by_descending(
                chunks, current, line,
            )
        }
        "dotnet.linq_select" => {
            crate::emitter::core::linq_adapter::emit_linq_select(chunks, current, line)
        }
        "dotnet.linq_select_many" => {
            crate::emitter::core::linq_adapter::emit_linq_select_many(chunks, current, line)
        }
        "dotnet.linq_group_by" => {
            crate::emitter::core::linq_adapter::emit_linq_group_by(chunks, current, line)
        }
        "dotnet.linq_to_dictionary" => {
            crate::emitter::core::linq_adapter::emit_linq_to_dictionary(
                chunks, current, line,
            )
        }
        "dotnet.linq_zip" => {
            crate::emitter::core::linq_adapter::emit_linq_zip(chunks, current, line)
        }
        "dotnet.linq_element_at" => {
            crate::emitter::core::linq_adapter::emit_linq_element_at(chunks, current, line)
        }
        "dotnet.linq_element_at_or_default" => {
            crate::emitter::core::linq_adapter::emit_linq_element_at_or_default(
                chunks, current, line,
            )
        }
        "dotnet.linq_single" => {
            crate::emitter::core::linq_adapter::emit_linq_single(chunks, current, line)
        }
        "dotnet.linq_single_or_default" => {
            crate::emitter::core::linq_adapter::emit_linq_single_or_default(chunks, current, line)
        }
        "dotnet.linq_max_by" => {
            crate::emitter::core::linq_adapter::emit_linq_max_by(chunks, current, line)
        }
        "dotnet.linq_min_by" => {
            crate::emitter::core::linq_adapter::emit_linq_min_by(chunks, current, line)
        }
        "dotnet.linq_aggregate_no_seed" => {
            crate::emitter::core::linq_adapter::emit_linq_aggregate_no_seed(chunks, current, line)
        }
        "dotnet.linq_append" => {
            crate::emitter::core::linq_adapter::emit_linq_append(chunks, current, line)
        }
        "dotnet.linq_prepend" => {
            crate::emitter::core::linq_adapter::emit_linq_prepend(chunks, current, line)
        }
        "dotnet.linq_sum" => {
            crate::emitter::core::linq_adapter::emit_linq_sum(chunks, current, line)
        }
        "dotnet.linq_count" => {
            crate::emitter::core::linq_adapter::emit_linq_count(chunks, current, line)
        }
        "dotnet.linq_skip_last" => {
            crate::emitter::core::linq_adapter::emit_linq_skip_last(chunks, current, line)
        }
        "dotnet.linq_take_last" => {
            crate::emitter::core::linq_adapter::emit_linq_take_last(chunks, current, line)
        }
        "dotnet.linq_default_if_empty" => {
            crate::emitter::core::linq_adapter::emit_linq_default_if_empty(chunks, current, line)
        }

        // ── Static Array.* helpers — same dotnet/core home as LINQ ──
        "dotnet.array_reverse" => {
            crate::emitter::core::array_adapter::emit_array_reverse(chunks, current, line)
        }
        "dotnet.array_index_of" => {
            crate::emitter::core::array_adapter::emit_array_index_of(chunks, current, line)
        }
        "dotnet.array_exists" => {
            crate::emitter::core::array_adapter::emit_array_exists(chunks, current, line)
        }
        "dotnet.array_true_for_all" => {
            crate::emitter::core::array_adapter::emit_array_true_for_all(
                chunks, current, line,
            )
        }
        "dotnet.array_find" => {
            crate::emitter::core::array_adapter::emit_array_find(chunks, current, line)
        }
        "dotnet.array_find_all" => {
            crate::emitter::core::array_adapter::emit_array_find_all(chunks, current, line)
        }
        "dotnet.array_convert_all" => {
            crate::emitter::core::array_adapter::emit_array_convert_all(
                chunks, current, line,
            )
        }
        "dotnet.array_for_each" => {
            crate::emitter::core::array_adapter::emit_array_for_each(chunks, current, line)
        }
        "dotnet.list_add_range" => {
            crate::emitter::core::array_adapter::emit_list_add_range(chunks, current, line)
        }

        // ── .NET parse helpers — `int.Parse`, `double.Parse`, `bool.Parse`
        // Throw a `FormatException`-shape error on invalid input
        // (matches ECMA-335; JS `Number(s)` returns NaN silently).
        "dotnet.parse_int" => {
            crate::emitter::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.parse_byte" => {
            crate::emitter::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.parse_long" => {
            crate::emitter::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.parse_float" => {
            crate::emitter::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_decimal" => {
            crate::emitter::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_double" => {
            crate::emitter::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_bool" => {
            crate::emitter::core::parse_adapter::emit_parse_bool(chunks, current, line)
        }
        "dotnet.parse_char" => {
            crate::emitter::core::parse_adapter::emit_parse_char(chunks, current, line)
        }

        // ── .NET System.Data adapter ────────────────────────────────
        "dotnet.datatable_new" => {
            crate::emitter::core::datatable_adapter::emit_datatable_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.dataset_new" => crate::emitter::core::datatable_adapter::emit_dataset_new(
            chunks, current, argc, line,
        ),
        "dotnet.datarow_new" => crate::emitter::core::datatable_adapter::emit_datarow_new(
            &mut chunks[current],
            line,
        ),
        "dotnet.datatable_new_row" => {
            crate::emitter::core::datatable_adapter::emit_datatable_new_row(
                chunks, current, line,
            )
        }
        "dotnet.datatable_add_row" => {
            crate::emitter::core::datatable_adapter::emit_datatable_add_row(
                chunks, current, line,
            )
        }
        "dotnet.datatable_select" => {
            crate::emitter::core::datatable_adapter::emit_datatable_select(
                chunks, current, line,
            )
        }
        "dotnet.dataset_tables" => {
            crate::emitter::core::datatable_adapter::emit_dataset_tables(
                chunks, current, line,
            )
        }
        "dotnet.datarow_item" => {
            crate::emitter::core::datatable_adapter::emit_datarow_item(
                chunks, current, line,
            )
        }
        "dotnet.datarow_is_null" => {
            crate::emitter::core::datatable_adapter::emit_datarow_is_null(
                chunks, current, line,
            )
        }

        // ── PHP `isset(...)` — variadic null check, returns true iff
        // ALL args are non-null. Inline emit folds an AND chain.
        "dotnet.dict_get_or_throw" => {
            // map[key] — get or throw KeyNotFoundException
            let chunk = &mut chunks[current];
            let has = chunk.add_import("ecma:map", "has");
            let get = chunk.add_import("ecma:map", "get");
            let key_slot = chunk.alloc_scratch(1);
            let map_slot = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
            chunk.emit_call(has, 2, line);
            chunk.emit_if_value(line);
            chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
            chunk.emit_call(get, 2, line);
            chunk.emit_else(line);
            chunk.emit_string_const("The given key was not present in the dictionary.", line);
            vybe_emitter::errors::emit_exception_new_finalize(
                chunk,
                "KeyNotFoundException",
                line,
            );
            vybe_emitter::errors::emit_throw(chunk, line);
            chunk.emit_end(line);
        }
        "dotnet.dict_get_value_or_default" => {
            // Stack: [map, key] or [map, key, default]
            // map.has(key) ? map.get(key) : default
            let chunk = &mut chunks[current];
            let has = chunk.add_import("ecma:map", "has");
            let get = chunk.add_import("ecma:map", "get");
            if argc >= 3 {
                // Explicit default: [map, key, default]
                let default_slot = chunk.alloc_scratch(1);
                let key_slot = chunk.alloc_scratch(1);
                let map_slot = chunk.alloc_scratch(1);
                chunk.emit_op_u16(Op::LOCAL_SET, default_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
                chunk.emit_call(has, 2, line);
                chunk.emit_if_value(line);
                chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
                chunk.emit_call(get, 2, line);
                chunk.emit_else(line);
                chunk.emit_op_u16(Op::LOCAL_GET, default_slot, line);
                chunk.emit_end(line);
            } else {
                // No explicit default: [map, key] → default is 0
                let key_slot = chunk.alloc_scratch(1);
                let map_slot = chunk.alloc_scratch(1);
                chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
                chunk.emit_call(has, 2, line);
                chunk.emit_if_value(line);
                chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
                chunk.emit_call(get, 2, line);
                chunk.emit_else(line);
                chunk.emit_f64_const(0.0, line);
                chunk.emit_end(line);
            }
        }
        "dotnet.dict_try_get_value" => {
            // TryGetValue(key, out value) → has(key) ? (value=get(key), true) : (value=default, false)
            // Simplified: returns get(key) or null, caller checks
            let chunk = &mut chunks[current];
            let out_slot = chunk.alloc_scratch(1);
            let key_slot = chunk.alloc_scratch(1);
            let map_slot = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
            let has = chunk.add_import("ecma:map", "has");
            chunk.emit_call(has, 2, line);
            chunk.emit_if_value(line);
            chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
            let get = chunk.add_import("ecma:map", "get");
            chunk.emit_call(get, 2, line);
            chunk.emit_else(line);
            chunk.emit_f64_const(0.0, line);
            chunk.emit_end(line);
        }
        "dotnet.choose" => emit_choose(&mut chunks[current], argc, line),

        // ── System.Math — shared .NET BCL math surface ──────────────
        // WASM opcodes (zero overhead)
        "dotnet.system.math.abs" => chunks[current].emit_op(Op::F64_ABS, line),
        "dotnet.system.math.floor" => chunks[current].emit_op(Op::F64_FLOOR, line),
        "dotnet.system.math.ceiling" | "dotnet.system.math.ceil" => {
            chunks[current].emit_op(Op::F64_CEIL, line)
        }
        "dotnet.system.math.sqrt" => chunks[current].emit_op(Op::F64_SQRT, line),
        "dotnet.system.math.truncate" | "dotnet.system.math.trunc" => {
            chunks[current].emit_op(Op::F64_TRUNC, line)
        }
        "dotnet.system.math.round" => {
            if argc <= 1 {
                chunks[current].emit_op(Op::F64_NEAREST, line);
            } else {
                // Round(value, digits): nearest(value * 10^digits) / 10^digits
                let chunk = &mut chunks[current];
                let digits_slot = chunk.alloc_scratch(1);
                let val_slot = chunk.alloc_scratch(1);
                let factor_slot = chunk.alloc_scratch(1);
                chunk.emit_op_u16(Op::LOCAL_SET, digits_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, val_slot, line);
                chunk.emit_f64_const(10.0, line);
                chunk.emit_op_u16(Op::LOCAL_GET, digits_slot, line);
                let pow = chunk.add_import("ecma:math", "pow");
                chunk.emit_call(pow, 2, line);
                chunk.emit_op_u16(Op::LOCAL_SET, factor_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, factor_slot, line);
                chunk.emit_op(Op::F64_MUL, line);
                chunk.emit_op(Op::F64_NEAREST, line);
                chunk.emit_op_u16(Op::LOCAL_GET, factor_slot, line);
                chunk.emit_op(Op::F64_DIV, line);
            }
        }
        "dotnet.system.math.min" => chunks[current].emit_op(Op::F64_MIN, line),
        "dotnet.system.math.max" => chunks[current].emit_op(Op::F64_MAX, line),
        // Host calls (ecma:math)
        "dotnet.system.math.pow" => {
            let idx = chunks[current].add_import("ecma:math", "pow");
            chunks[current].emit_call(idx, 2, line);
        }
        "dotnet.system.math.sin" => {
            let idx = chunks[current].add_import("ecma:math", "sin");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.cos" => {
            let idx = chunks[current].add_import("ecma:math", "cos");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.tan" => {
            let idx = chunks[current].add_import("ecma:math", "tan");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.asin" => {
            let idx = chunks[current].add_import("ecma:math", "asin");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.acos" => {
            let idx = chunks[current].add_import("ecma:math", "acos");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.atan" => {
            let idx = chunks[current].add_import("ecma:math", "atan");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.atan2" => {
            let idx = chunks[current].add_import("ecma:math", "atan2");
            chunks[current].emit_call(idx, 2, line);
        }
        "dotnet.system.math.log" => {
            let idx = chunks[current].add_import("ecma:math", "log");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.log10" => {
            let idx = chunks[current].add_import("ecma:math", "log10");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.log2" => {
            let idx = chunks[current].add_import("ecma:math", "log2");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.exp" => {
            let idx = chunks[current].add_import("ecma:math", "exp");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.sinh" => {
            let idx = chunks[current].add_import("ecma:math", "sinh");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.cosh" => {
            let idx = chunks[current].add_import("ecma:math", "cosh");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.tanh" => {
            let idx = chunks[current].add_import("ecma:math", "tanh");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.sign" => {
            let idx = chunks[current].add_import("ecma:math", "sign");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.clamp" => vybe_emitter::math::emit_clamp(&mut chunks[current], line),
        _ => return false,
    }
    true
}
