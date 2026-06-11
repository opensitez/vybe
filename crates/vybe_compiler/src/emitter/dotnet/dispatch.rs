//! Auto-extracted `dotnet.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

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
    let arr_slot = chunk.local_count;
    let idx_slot = arr_slot + 1;
    chunk.local_count = arr_slot + 2;

    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, n, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_op(Op::I32_CONST_1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if crate::emitter::dotnet::core::runtime_adapter::emit_helper(name, chunks, current, argc, line)
    {
        return true;
    }
    match name {
        "dotnet.dns_get_host_addresses" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_dns_get_host_addresses(
                chunks, current, line,
            )
        }
        "dotnet.dns_get_host_entry" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_dns_get_host_entry(
                chunks, current, line,
            )
        }
        "dotnet.dns_get_host_name" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_dns_get_host_name(
                chunks, current, line,
            )
        }
        "dotnet.tcp_client_new" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_tcp_client_new(
                chunks, current, line,
            )
        }
        "dotnet.tcp_client_get_stream" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_tcp_client_get_stream(
                chunks, current, line,
            )
        }
        "dotnet.tcp_client_close" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_tcp_client_close(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_new" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_new(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_start" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_start(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_stop" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_stop(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_accept" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_accept(
                chunks, current, line,
            )
        }
        "dotnet.tcp_listener_pending" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_pending(
                chunks, current, line,
            )
        }
        "dotnet.udp_client_new" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_udp_client_new(
                chunks, current, line,
            )
        }
        "dotnet.udp_send" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_udp_send(chunks, current, line)
        }
        "dotnet.udp_receive" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_udp_receive(chunks, current, line)
        }
        "dotnet.udp_close" => {
            crate::emitter::dotnet::core::sockets_adapter::emit_udp_close(chunks, current, line)
        }

        // ── .NET StringBuilder adapter ──────────────────────────────
        // No direct ECMA mirror; the wrapper materializes a plain
        // Object with a `__buffer` string and mutates via DYN_ADD +
        // STRUCT_SET. Multi-arity ctor uses the threaded `argc` to
        // pick between empty / initial-keyed shapes.
        "dotnet.string_builder_new" => {
            crate::emitter::dotnet::core::stringbuilder_adapter::emit_string_builder_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.sb_append" => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_append(
            chunks, current, line,
        ),
        "dotnet.sb_append_line" => {
            crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_append_line(
                chunks, current, line,
            )
        }
        "dotnet.sb_to_string" => {
            crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_to_string(
                chunks, current, line,
            )
        }
        "dotnet.sb_clear" => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_clear(
            chunks, current, line,
        ),
        "dotnet.sb_length" => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_length(
            chunks, current, line,
        ),
        "dotnet.sb_insert" => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_insert(
            chunks, current, line,
        ),
        "dotnet.sb_remove" => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_remove(
            chunks, current, line,
        ),
        "dotnet.sb_replace" => {
            crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_replace(
                chunks, current, line,
            )
        }

        // ── .NET Random adapter ─────────────────────────────────────
        "dotnet.random_new" => crate::emitter::dotnet::core::random_adapter::emit_random_new(
            chunks, current, argc, line,
        ),
        "dotnet.random_next" => crate::emitter::dotnet::core::random_adapter::emit_random_next(
            chunks, current, argc, line,
        ),
        "dotnet.random_next_double" => {
            crate::emitter::dotnet::core::random_adapter::emit_random_next_double(
                chunks, current, line,
            )
        }

        // ── .NET Regex adapter ──────────────────────────────────────
        "dotnet.regex_new" => {
            crate::emitter::dotnet::core::regex_adapter::emit_regex_new(chunks, current, argc, line)
        }
        "dotnet.regex_is_match" => {
            crate::emitter::dotnet::core::regex_adapter::emit_regex_is_match(chunks, current, line)
        }
        "dotnet.regex_replace" => {
            crate::emitter::dotnet::core::regex_adapter::emit_regex_replace(chunks, current, line)
        }
        "dotnet.regex_split" => {
            crate::emitter::dotnet::core::regex_adapter::emit_regex_split(chunks, current, line)
        }
        "dotnet.regex_match" => {
            crate::emitter::dotnet::core::regex_adapter::emit_regex_match(chunks, current, line)
        }
        "dotnet.regex_matches" => {
            crate::emitter::dotnet::core::regex_adapter::emit_regex_matches(chunks, current, line)
        }

        // ── .NET Stopwatch adapter ──────────────────────────────────
        "dotnet.stopwatch_new" => {
            crate::emitter::dotnet::core::stopwatch_adapter::emit_stopwatch_new(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_start" => {
            crate::emitter::dotnet::core::stopwatch_adapter::emit_stopwatch_start(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_stop" => {
            crate::emitter::dotnet::core::stopwatch_adapter::emit_stopwatch_stop(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_reset" => {
            crate::emitter::dotnet::core::stopwatch_adapter::emit_stopwatch_reset(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_start_new" => {
            crate::emitter::dotnet::core::stopwatch_adapter::emit_stopwatch_start_new(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_restart" => {
            crate::emitter::dotnet::core::stopwatch_adapter::emit_stopwatch_restart(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_elapsed_ms" => {
            crate::emitter::dotnet::core::stopwatch_adapter::emit_stopwatch_elapsed_ms(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_is_running" => {
            crate::emitter::dotnet::core::stopwatch_adapter::emit_stopwatch_is_running(
                chunks, current, line,
            )
        }

        // ── .NET Process / ProcessStartInfo adapter ─────────────────
        // Lowers to `node:child_process.spawnSync` + plain Object
        // structs for the .NET-shape records. Multi-arity ctors use
        // the threaded `argc`.
        "dotnet.process_start_info_new" => {
            crate::emitter::dotnet::core::process_adapter::emit_process_start_info_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.process_new" => crate::emitter::dotnet::core::process_adapter::emit_process_new(
            chunks, current, argc, line,
        ),
        "dotnet.process_start" => {
            crate::emitter::dotnet::core::process_adapter::emit_process_start(chunks, current, line)
        }
        "dotnet.process_get_current" => {
            crate::emitter::dotnet::core::process_adapter::emit_process_get_current(
                chunks, current, line,
            )
        }
        "dotnet.process_wait_for_exit" => {
            crate::emitter::dotnet::core::process_adapter::emit_process_wait_for_exit(
                chunks, current, line,
            )
        }

        // ── .NET System.Array static-method adapter ─────────────────
        // `Clear` / `Copy` / `Resize` / `Sort` lower to bundled stdlib
        // chunks (`__vybe_*` globals) composing `ecma:array.*`
        // primitives. No `vybe:types/array*` host fns.
        "dotnet.array_clear" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_clear(chunks, current, line)
        }
        "dotnet.array_copy" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_copy(chunks, current, line)
        }
        "dotnet.array_resize" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_resize(chunks, current, line)
        }
        "dotnet.array_sort" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_sort(chunks, current, line)
        }
        "dotnet.hashset_add" => {
            crate::emitter::dotnet::core::collections_adapter::emit_hashset_add(
                chunks, current, line,
            )
        }
        "dotnet.hashset_union_with" => {
            crate::emitter::dotnet::core::collections_adapter::emit_hashset_union_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_intersect_with" => {
            crate::emitter::dotnet::core::collections_adapter::emit_hashset_intersect_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_except_with" => {
            crate::emitter::dotnet::core::collections_adapter::emit_hashset_except_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_symmetric_except_with" => {
            crate::emitter::dotnet::core::collections_adapter::emit_hashset_symmetric_except_with(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_add_first" => {
            crate::emitter::dotnet::core::collections_adapter::emit_linked_list_add_first(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_add_last" => {
            crate::emitter::dotnet::core::collections_adapter::emit_linked_list_add_last(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_find" => {
            crate::emitter::dotnet::core::collections_adapter::emit_linked_list_find(
                chunks, current, line,
            )
        }
        "dotnet.sorted_dictionary_entries" => {
            crate::emitter::dotnet::core::collections_adapter::emit_sorted_dictionary_entries(
                chunks, current, line,
            )
        }

        // ── .NET TimeSpan factory adapters ──────────────────────────
        // `TimeSpan.From*(n)` factories build a duration record by
        // multiplying `n` with the unit-to-ms factor. Pure inline
        // bytecode; no host fns.
        "dotnet.timespan_from_days" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_days(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_hours" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_hours(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_minutes" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_minutes(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_seconds" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_seconds(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_milliseconds" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_milliseconds(
                chunks, current, line,
            )
        }
        "dotnet.timespan_zero" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_zero(
                chunks, current, line,
            )
        }
        "dotnet.timespan_new" => crate::emitter::dotnet::core::timespan_adapter::emit_timespan_new(
            chunks, current, argc, line,
        ),
        "dotnet.timespan_compare" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_compare(
                chunks, current, line,
            )
        }
        "dotnet.timespan_negate" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_negate(
                chunks, current, line,
            )
        }
        "dotnet.timespan_duration" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_duration(
                chunks, current, line,
            )
        }
        "dotnet.timespan_add" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_add(chunks, current, line)
        }
        "dotnet.timespan_sub" => {
            crate::emitter::dotnet::core::timespan_adapter::emit_timespan_sub(chunks, current, line)
        }

        // ── .NET Guid adapters ──────────────────────────────────────
        // `Guid` is stored as a .NET-shaped object carrying the
        // canonical lowercase text representation in `__value`.
        "dotnet.guid_empty" => {
            crate::emitter::dotnet::core::guid_adapter::emit_guid_empty(chunks, current, line)
        }
        "dotnet.guid_new_guid" => {
            crate::emitter::dotnet::core::guid_adapter::emit_guid_new_guid(chunks, current, line)
        }
        "dotnet.guid_parse" => {
            crate::emitter::dotnet::core::guid_adapter::emit_guid_parse(chunks, current, line)
        }
        "dotnet.guid_new" => {
            crate::emitter::dotnet::core::guid_adapter::emit_guid_new(chunks, current, argc, line)
        }
        "dotnet.guid_to_string" => {
            crate::emitter::dotnet::core::guid_adapter::emit_guid_to_string(chunks, current, line)
        }
        "dotnet.guid_try_parse" => crate::emitter::dotnet::core::guid_adapter::emit_guid_try_parse(
            chunks, current, argc, line,
        ),

        "dotnet.version_new" => crate::emitter::dotnet::core::version_adapter::emit_version_new(
            chunks, current, argc, line,
        ),
        "dotnet.version_parse" => {
            crate::emitter::dotnet::core::version_adapter::emit_version_parse(chunks, current, line)
        }
        "dotnet.version_to_string" => {
            crate::emitter::dotnet::core::version_adapter::emit_version_to_string(
                chunks, current, line,
            )
        }
        "dotnet.version_compare" => {
            crate::emitter::dotnet::core::version_adapter::emit_version_compare(
                chunks, current, line,
            )
        }
        "dotnet.version_equals" => {
            crate::emitter::dotnet::core::version_adapter::emit_version_equals(
                chunks, current, line,
            )
        }
        "dotnet.version_lt" => {
            crate::emitter::dotnet::core::version_adapter::emit_version_lt(chunks, current, line)
        }
        "dotnet.version_gt" => {
            crate::emitter::dotnet::core::version_adapter::emit_version_gt(chunks, current, line)
        }
        "dotnet.version_eq" => {
            crate::emitter::dotnet::core::version_adapter::emit_version_eq(chunks, current, line)
        }
        "dotnet.version_ne" => {
            crate::emitter::dotnet::core::version_adapter::emit_version_ne(chunks, current, line)
        }

        // ── .NET DateTime static adapters ───────────────────────────
        // `Now` / `UtcNow` / `Today` lower to `ecma:date.now` (which
        // reads `wasi:clocks/wall-clock.now`); `Parse` lowers to
        // `ecma:date.parse`. Each wraps the resulting ms timestamp
        // in a `{__type:"DateTime", __time:ms}` object so the .NET
        // surface looks .NET-shaped.
        "dotnet.datetime_now" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_now(chunks, current, line)
        }
        "dotnet.datetime_parse" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_parse(
                chunks, current, line,
            )
        }
        "dotnet.datetime_today" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_today(
                chunks, current, line,
            )
        }
        "dotnet.datetime_new" => crate::emitter::dotnet::core::datetime_adapter::emit_datetime_new(
            chunks, current, argc, line,
        ),
        "dotnet.datetime_add_days" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_add_days(
                chunks, current, line,
            )
        }
        "dotnet.datetime_add_hours" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_add_hours(
                chunks, current, line,
            )
        }
        "dotnet.datetime_add_months" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_add_months(
                chunks, current, line,
            )
        }
        "dotnet.datetime_days_in_month" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_days_in_month(
                chunks, current, line,
            )
        }
        "dotnet.datetime_is_leap_year" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_is_leap_year(
                chunks, current, line,
            )
        }
        "dotnet.datetime_compare" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_compare(
                chunks, current, line,
            )
        }
        "dotnet.datetime_to_short_date_string" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_to_short_date_string(
                chunks, current, line,
            )
        }
        "dotnet.datetime_add_timespan" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_add_timespan(
                chunks, current, line,
            )
        }
        "dotnet.datetime_subtract_datetime" => {
            crate::emitter::dotnet::core::datetime_adapter::emit_datetime_subtract_datetime(
                chunks, current, line,
            )
        }

        // ── PHP DateTime / DateTimeImmutable / DateInterval adapters ──
        // Bytecode-only — composes existing `ecma:date.*` host fns into
        // the PHP-shaped surface. See `emitter/php/datetime_adapter.rs`.
        "dotnet.string_format" => {
            crate::emitter::dotnet::core::string_format_adapter::emit_string_format(
                chunks, current, argc, line,
            )
        }

        // ── VB / VBA `Format(value, picture)` — picture-string render ──
        "dotnet.format_picture" => {
            crate::emitter::dotnet::core::format_picture_adapter::emit_format_picture(
                chunks, current, argc, line,
            )
        }

        // ── .NET StreamReader / StreamWriter adapters — text I/O ────
        // Load-whole-file model: `new StreamReader(path)` materializes a
        // string buffer via `node:fs.readFileSync`, `new StreamWriter`
        // accumulates into `__buf` and flushes via `writeFileSync`.
        // Bytecode-only — no `dotnet:io` host fns.
        "dotnet.stream_reader_new" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.stream_reader_read_line" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_read_line(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_read_to_end" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_read_to_end(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_at_end" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_at_end(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_close" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_close(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_new" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_writer_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.stream_writer_write" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_writer_write(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_write_line" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_writer_write_line(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_flush" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_writer_flush(
                chunks, current, line,
            )
        }
        "dotnet.stream_close" => {
            crate::emitter::dotnet::core::stream_io_adapter::emit_stream_close(
                chunks, current, line,
            )
        }
        "dotnet.file_read_all_lines" => {
            crate::emitter::dotnet::core::filesystem_adapter::emit_file_read_all_lines(
                chunks, current, line,
            )
        }
        "dotnet.directory_get_files" => {
            crate::emitter::dotnet::core::filesystem_adapter::emit_directory_get_files(
                chunks, current, line,
            )
        }
        "dotnet.directory_get_directories" => {
            crate::emitter::dotnet::core::filesystem_adapter::emit_directory_get_directories(
                chunks, current, line,
            )
        }
        "dotnet.console_writeline" => {
            crate::emitter::dotnet::core::console_adapter::emit_console_writeline(
                chunks, current, line,
            )
        }
        "dotnet.console_readline" => {
            crate::emitter::dotnet::core::console_adapter::emit_console_readline(
                chunks, current, line,
            )
        }
        "dotnet.console_error" => {
            crate::emitter::dotnet::core::console_adapter::emit_console_error(
                chunks, current, line,
            )
        }
        "dotnet.environment_username" => {
            crate::emitter::dotnet::core::environment_adapter::emit_environment_username(
                chunks, current, line,
            )
        }
        "dotnet.environment_processor_count" => {
            crate::emitter::dotnet::core::environment_adapter::emit_environment_processor_count(
                chunks, current, line,
            )
        }
        "dotnet.environment_tick_count" => {
            crate::emitter::dotnet::core::environment_adapter::emit_environment_tick_count(
                chunks, current, line,
            )
        }
        "dotnet.environment_get" => {
            crate::emitter::dotnet::core::environment_adapter::emit_environment_get(
                chunks, current, line,
            )
        }
        "dotnet.environment_set" => {
            crate::emitter::dotnet::core::environment_adapter::emit_environment_set(
                chunks, current, line,
            )
        }

        // ── OleDb adapter — System.Data.OleDb constructor wrappers ─────────────
        "dotnet.oledb_connection_new" => {
            crate::emitter::dotnet::core::oledb_adapter::emit_oledb_connection_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.oledb_command_new" => {
            crate::emitter::dotnet::core::oledb_adapter::emit_oledb_command_new(
                chunks, current, argc, line,
            )
        }

        // ── ADODB adapter — ADODB.Connection / Command / Recordset ──────────────
        "dotnet.adodb_connection_new" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_connection_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_connection_execute" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_connection_execute(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_begin_trans" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_conn_begin_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_commit_trans" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_conn_commit_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_rollback_trans" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_conn_rollback_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_command_new" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_command_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_command_execute" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_command_execute(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_command_create_parameter" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_command_create_parameter(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_new" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_recordset_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_open" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_recordset_open(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_move_next" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_recordset_move_next(
                chunks, current, line,
            )
        }
        "dotnet.adodb_recordset_move_first" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_recordset_move_first(
                chunks, current, line,
            )
        }
        "dotnet.adodb_recordset_fields" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_recordset_fields(
                chunks, current, line,
            )
        }
        "dotnet.adodb_recordset_close" => {
            crate::emitter::dotnet::core::adodb_adapter::emit_adodb_recordset_close(
                chunks, current, line,
            )
        }

        // ── LINQ surface — composed bytecode shared by every .NET-shape language ──
        "dotnet.linq_first" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_first(chunks, current, line)
        }
        "dotnet.linq_last" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_last(chunks, current, line)
        }
        "dotnet.linq_skip" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_skip(chunks, current, line)
        }
        "dotnet.linq_take" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_take(chunks, current, line)
        }
        "dotnet.linq_identity" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_identity(chunks, current, line)
        }
        "dotnet.linq_average" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_average(chunks, current, line)
        }
        "dotnet.linq_first_or_default" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_first_or_default(
                chunks, current, line,
            )
        }
        "dotnet.linq_distinct" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_distinct(chunks, current, line)
        }
        "dotnet.linq_sequence_equal" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_sequence_equal(
                chunks, current, line,
            )
        }
        "dotnet.linq_count_pred" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_count_pred(chunks, current, line)
        }
        "dotnet.linq_aggregate" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_aggregate(chunks, current, line)
        }
        "dotnet.linq_order_by_descending" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_order_by_descending(
                chunks, current, line,
            )
        }
        "dotnet.linq_select" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_select(chunks, current, line)
        }
        "dotnet.linq_select_many" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_select_many(chunks, current, line)
        }
        "dotnet.linq_group_by" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_group_by(chunks, current, line)
        }
        "dotnet.linq_to_dictionary" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_to_dictionary(
                chunks, current, line,
            )
        }
        "dotnet.linq_zip" => {
            crate::emitter::dotnet::core::linq_adapter::emit_linq_zip(chunks, current, line)
        }

        // ── Static Array.* helpers — same dotnet/core home as LINQ ──
        "dotnet.array_reverse" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_reverse(chunks, current, line)
        }
        "dotnet.array_index_of" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_index_of(chunks, current, line)
        }
        "dotnet.array_exists" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_exists(chunks, current, line)
        }
        "dotnet.array_true_for_all" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_true_for_all(
                chunks, current, line,
            )
        }
        "dotnet.array_find" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_find(chunks, current, line)
        }
        "dotnet.array_find_all" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_find_all(chunks, current, line)
        }
        "dotnet.array_convert_all" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_convert_all(
                chunks, current, line,
            )
        }
        "dotnet.array_for_each" => {
            crate::emitter::dotnet::core::array_adapter::emit_array_for_each(chunks, current, line)
        }
        "dotnet.list_add_range" => {
            crate::emitter::dotnet::core::array_adapter::emit_list_add_range(chunks, current, line)
        }

        // ── .NET parse helpers — `int.Parse`, `double.Parse`, `bool.Parse`
        // Throw a `FormatException`-shape error on invalid input
        // (matches ECMA-335; JS `Number(s)` returns NaN silently).
        "dotnet.parse_int" => {
            crate::emitter::dotnet::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.parse_byte" => {
            crate::emitter::dotnet::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.parse_long" => {
            crate::emitter::dotnet::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.parse_float" => {
            crate::emitter::dotnet::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_decimal" => {
            crate::emitter::dotnet::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_double" => {
            crate::emitter::dotnet::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_bool" => {
            crate::emitter::dotnet::core::parse_adapter::emit_parse_bool(chunks, current, line)
        }
        "dotnet.parse_char" => {
            crate::emitter::dotnet::core::parse_adapter::emit_parse_char(chunks, current, line)
        }

        // ── PHP `isset(...)` — variadic null check, returns true iff
        // ALL args are non-null. Inline emit folds an AND chain.
        "dotnet.choose" => emit_choose(&mut chunks[current], argc, line),
        _ => return false,
    }
    true
}
