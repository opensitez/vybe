//! Centralized `common:<name>` emit dispatcher.
//!
//! Language profiles use the `emit = "common:<category>.<op>"` convention to
//! delegate to a canonical compiler_common helper. This module owns the
//! `<name> → emit fn` mapping so every language compiler shares one source
//! of truth — adding a new common op only needs to be done here, and every
//! frontend that uses the dispatcher gets it for free.
//!
//! ## Two flavors
//!
//! - `emit_common(name, chunk, line)` handles ops that need ONLY a chunk and
//!   line (the vast majority — pure bytecode emits).
//! - `emit_common_with_imports(name, chunk, line, import)` handles ops that
//!   ALSO need to register a host import (e.g. `threading.sleep` adds a
//!   `wasi:clocks::sleep` import). The `import` callback resolves the import
//!   index in whatever way the host compiler does (typically by adding to a
//!   designated chunk's import table).
//!
//! Both functions return `true` if they recognized and emitted `name`, and
//! `false` if the name is unknown — letting the caller fall through to its
//! own dispatch for language-specific common ops.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

use crate::emitter::{collections, dict, strings, threading};

/// Handle common ops that need only a chunk and line.
/// Returns `true` if `name` was recognized and emitted, `false` otherwise.
///
/// `argc` is the number of caller-supplied values currently on the
/// stack at the emit site. Most emits ignore it (their stack contract
/// is fixed — `dict.has` always pops two), but multi-arity emits like
/// .NET constructors with overloaded shapes (`new StringBuilder()` vs
/// `new StringBuilder("initial")`) branch on it to pick the right
/// bytecode.
///
/// Takes `&mut Vec<Chunk>` rather than `&mut [Chunk]` because some helpers
/// (e.g. `threading.task_delay`) push a new function chunk for the worker
/// body. Slice-shape ops still work — `&mut Vec<Chunk>` derefs to
/// `&mut [Chunk]` for index access.
pub fn emit_common(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        // ── Dict ops ──
        "dict.set_dynamic" => {
            dict::emit_set_dynamic(chunks, current, line);
            chunks[current].emit_op(Op::NULL, line); // void return
        }
        "dict.get_dynamic" => dict::emit_get_dynamic(chunks, current, line),
        "dict.has" => dict::emit_method_has(chunks, current, line),
        "dict.delete" => dict::emit_method_delete(chunks, current, line),
        "dict.clear" => dict::emit_method_clear_stack(chunks, current, line),
        "dict.size" => dict::emit_method_size(chunks, current, line),
        "dict.keys" => dict::emit_keys(chunks, current, line),
        "dict.values" => dict::emit_values(chunks, current, line),
        "dict.new" => dict::emit_new(chunks, current, line),

        // ── Object ops ── ecma:object/new creates a plain JS Object.
        // The `.NET` Dictionary class uses this as its backing (matches
        // the ECMA-262 rule that a Dictionary<string, T> is shape-identical
        // to an Object). Method dispatch routes through `ecma:object/*`
        // via TypeRegistry, so no parallel vybe:types/dict* host fns are
        // consulted.
        "object.new" => {
            let idx = chunks[0].add_import("ecma:object", "new");
            chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
            chunks[current].emit(0, line);
        }

        // ── Collection ops (route through ecma:array imports;
        // `chunks` slice lets the helper register on chunks[0] while
        // emitting code on chunks[current]). ──
        "collections.push" => collections::emit_push(chunks, current, line),
        "collections.pop" => collections::emit_pop(chunks, current, line),
        "collections.length" => collections::emit_len(chunks, current, line),
        "collections.get" => collections::emit_get(chunks, current, line),
        "collections.set" => collections::emit_set(chunks, current, line),
        "collections.contains" => collections::emit_contains(chunks, current, line),
        "collections.index_of" => collections::emit_index_of(chunks, current, line),
        "collections.last_index_of" => collections::emit_last_index_of(chunks, current, line),
        "collections.remove_at" => collections::emit_remove_at(chunks, current, line),
        "collections.sorted" => collections::emit_sorted(chunks, current, line),
        "collections.reverse" => collections::emit_reverse(chunks, current, line),
        "collections.join" => collections::emit_join(chunks, current, line),
        "collections.slice" => collections::emit_slice(chunks, current, line),
        "collections.new" => collections::emit_array_new(chunks, current, 0, line),
        "collections.shift" => collections::emit_shift(chunks, current, line),
        "collections.concat" => collections::emit_concat(chunks, current, line),
        "collections.fill" => collections::emit_fill(chunks, current, line),
        "collections.sort" => collections::emit_sort(chunks, current, line),
        "collections.index_of_from" => collections::emit_index_of_from(chunks, current, line),
        "collections.last_index_of_from" => collections::emit_last_index_of_from(chunks, current, line),
        "collections.remove_range" => collections::emit_remove_range(chunks, current, line),
        "collections.get_range" => collections::emit_get_range(chunks, current, line),
        "collections.clone" => collections::emit_clone(chunks, current, line),
        "collections.insert_range" => collections::emit_insert_range(chunks, current, line),
        "collections.set_range" => collections::emit_set_range(chunks, current, line),
        "collections.binary_search" => collections::emit_binary_search(chunks, current, line),
        "collections.reverse_range" => collections::emit_reverse_range(chunks, current, line),
        "collections.remove" => collections::emit_remove_value(chunks, current, line),
        "collections.insert" => collections::emit_insert_at(chunks, current, line),
        "collections.clear" => collections::emit_clear(chunks, current, line),

        // ── String ops ──
        "strings.length" => strings::emit_length(&mut chunks[current], line),
        "strings.to_upper" => strings::emit_to_upper(&mut chunks[current], line),
        "strings.to_lower" => strings::emit_to_lower(&mut chunks[current], line),
        "strings.trim" => strings::emit_trim(&mut chunks[current], line),
        "strings.substring" => strings::emit_substring(&mut chunks[current], line),
        "strings.replace" => strings::emit_replace(&mut chunks[current], line),
        "strings.split" => strings::emit_split(&mut chunks[current], line),
        "strings.index_of" => strings::emit_index_of(&mut chunks[current], line),
        "strings.concat" => strings::emit_concat(&mut chunks[current], 2, line),

        // ── Expression ops ──
        "expressions.undefined" => crate::emitter::expressions::emit_undefined(&mut chunks[current], line),
        "expressions.i32_not" => crate::emitter::expressions::emit_i32_not(&mut chunks[current], line),
        "expressions.f64_mod" => crate::emitter::expressions::emit_f64_mod(&mut chunks[current], line),
        "expressions.bool_not" => crate::emitter::expressions::emit_bool_not(&mut chunks[current], line),

        // ── Threading ops ──
        // Real WASM threading opcodes (wasi-threads proposal):
        // thread_spawn, thread_join, memory atomic_*. NOT host calls — these
        // run unchanged on any standard WASM runtime that supports the
        // threads proposal.
        "threading.task_run" => threading::emit_task_run(chunks, current, line),
        "threading.task_delay" => threading::emit_task_delay(chunks, current, line),
        "threading.thread_new" => threading::emit_thread_new(chunks, current, line),
        "threading.thread_spawn" => threading::emit_thread_spawn(chunks, current, line),

        // ── .NET sockets adapter ────────────────────────────────────
        // Each .NET socket method lowers to a sequence of
        // `wasi:sockets/*` / `wasi:io/streams.*` / `node:os.*` host
        // imports. The dotnet:* host modules retire — adapter logic
        // lives entirely in `emitter::dotnet::core::sockets_adapter`.
        "dotnet.dns_get_host_addresses"
            => crate::emitter::dotnet::core::sockets_adapter::emit_dns_get_host_addresses(chunks, current, line),
        "dotnet.dns_get_host_entry"
            => crate::emitter::dotnet::core::sockets_adapter::emit_dns_get_host_entry(chunks, current, line),
        "dotnet.dns_get_host_name"
            => crate::emitter::dotnet::core::sockets_adapter::emit_dns_get_host_name(chunks, current, line),
        "dotnet.tcp_client_new"
            => crate::emitter::dotnet::core::sockets_adapter::emit_tcp_client_new(chunks, current, line),
        "dotnet.tcp_client_get_stream"
            => crate::emitter::dotnet::core::sockets_adapter::emit_tcp_client_get_stream(chunks, current, line),
        "dotnet.tcp_client_close"
            => crate::emitter::dotnet::core::sockets_adapter::emit_tcp_client_close(chunks, current, line),
        "dotnet.tcp_listener_new"
            => crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_new(chunks, current, line),
        "dotnet.tcp_listener_start"
            => crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_start(chunks, current, line),
        "dotnet.tcp_listener_stop"
            => crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_stop(chunks, current, line),
        "dotnet.tcp_listener_accept"
            => crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_accept(chunks, current, line),
        "dotnet.tcp_listener_pending"
            => crate::emitter::dotnet::core::sockets_adapter::emit_tcp_listener_pending(chunks, current, line),
        "dotnet.udp_client_new"
            => crate::emitter::dotnet::core::sockets_adapter::emit_udp_client_new(chunks, current, line),
        "dotnet.udp_send"
            => crate::emitter::dotnet::core::sockets_adapter::emit_udp_send(chunks, current, line),
        "dotnet.udp_receive"
            => crate::emitter::dotnet::core::sockets_adapter::emit_udp_receive(chunks, current, line),
        "dotnet.udp_close"
            => crate::emitter::dotnet::core::sockets_adapter::emit_udp_close(chunks, current, line),

        // ── .NET StringBuilder adapter ──────────────────────────────
        // No direct ECMA mirror; the wrapper materializes a plain
        // Object with a `__buffer` string and mutates via DYN_ADD +
        // STRUCT_SET. Multi-arity ctor uses the threaded `argc` to
        // pick between empty / initial-keyed shapes.
        "dotnet.string_builder_new"
            => crate::emitter::dotnet::core::stringbuilder_adapter::emit_string_builder_new(chunks, current, argc, line),
        "dotnet.sb_append"
            => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_append(chunks, current, line),
        "dotnet.sb_append_line"
            => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_append_line(chunks, current, line),
        "dotnet.sb_to_string"
            => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_to_string(chunks, current, line),
        "dotnet.sb_clear"
            => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_clear(chunks, current, line),
        "dotnet.sb_length"
            => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_length(chunks, current, line),
        "dotnet.sb_insert"
            => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_insert(chunks, current, line),
        "dotnet.sb_replace"
            => crate::emitter::dotnet::core::stringbuilder_adapter::emit_sb_replace(chunks, current, line),

        // ── .NET Process / ProcessStartInfo adapter ─────────────────
        // Lowers to `node:child_process.spawnSync` + plain Object
        // structs for the .NET-shape records. Multi-arity ctors use
        // the threaded `argc`.
        "dotnet.process_start_info_new"
            => crate::emitter::dotnet::core::process_adapter::emit_process_start_info_new(chunks, current, argc, line),
        "dotnet.process_new"
            => crate::emitter::dotnet::core::process_adapter::emit_process_new(chunks, current, argc, line),
        "dotnet.process_start"
            => crate::emitter::dotnet::core::process_adapter::emit_process_start(chunks, current, line),
        "dotnet.process_get_current"
            => crate::emitter::dotnet::core::process_adapter::emit_process_get_current(chunks, current, line),
        "dotnet.process_wait_for_exit"
            => crate::emitter::dotnet::core::process_adapter::emit_process_wait_for_exit(chunks, current, line),

        // ── .NET System.Array static-method adapter ─────────────────
        // `Clear` / `Copy` / `Resize` / `Sort` lower to bundled stdlib
        // chunks (`__vybe_*` globals) composing `ecma:array.*`
        // primitives. No `vybe:types/array*` host fns.
        "dotnet.array_clear"
            => crate::emitter::dotnet::core::array_adapter::emit_array_clear(chunks, current, line),
        "dotnet.array_copy"
            => crate::emitter::dotnet::core::array_adapter::emit_array_copy(chunks, current, line),
        "dotnet.array_resize"
            => crate::emitter::dotnet::core::array_adapter::emit_array_resize(chunks, current, line),
        "dotnet.array_sort"
            => crate::emitter::dotnet::core::array_adapter::emit_array_sort(chunks, current, line),

        // ── .NET TimeSpan factory adapters ──────────────────────────
        // `TimeSpan.From*(n)` factories build a duration record by
        // multiplying `n` with the unit-to-ms factor. Pure inline
        // bytecode; no host fns.
        "dotnet.timespan_from_days"
            => crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_days(chunks, current, line),
        "dotnet.timespan_from_hours"
            => crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_hours(chunks, current, line),
        "dotnet.timespan_from_minutes"
            => crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_minutes(chunks, current, line),
        "dotnet.timespan_from_seconds"
            => crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_seconds(chunks, current, line),
        "dotnet.timespan_from_milliseconds"
            => crate::emitter::dotnet::core::timespan_adapter::emit_timespan_from_milliseconds(chunks, current, line),
        "dotnet.timespan_zero"
            => crate::emitter::dotnet::core::timespan_adapter::emit_timespan_zero(chunks, current, line),

        // ── .NET DateTime static adapters ───────────────────────────
        // `Now` / `UtcNow` / `Today` lower to `ecma:date.now` (which
        // reads `wasi:clocks/wall-clock.now`); `Parse` lowers to
        // `ecma:date.parse`. Each wraps the resulting ms timestamp
        // in a `{__type:"DateTime", __time:ms}` object so the .NET
        // surface looks .NET-shaped.
        "dotnet.datetime_now"
            => crate::emitter::dotnet::core::datetime_adapter::emit_datetime_now(chunks, current, line),
        "dotnet.datetime_parse"
            => crate::emitter::dotnet::core::datetime_adapter::emit_datetime_parse(chunks, current, line),
        "dotnet.datetime_today"
            => crate::emitter::dotnet::core::datetime_adapter::emit_datetime_today(chunks, current, line),

        // ── .NET String.Format adapter — composite-format substitution ──
        // `String.Format(fmt, ...args)` lowers to inline bytecode that
        // walks the format string, parses `{N}` / `{{` / `}}` tokens,
        // and concatenates. `argc` includes the format string; trailing
        // args are packed into a local array indexed by placeholder N.
        "dotnet.string_format"
            => crate::emitter::dotnet::core::string_format_adapter::emit_string_format(chunks, current, argc, line),

        // ── VB / VBA `Format(value, picture)` — picture-string render ──
        "dotnet.format_picture"
            => crate::emitter::dotnet::core::format_picture_adapter::emit_format_picture(chunks, current, argc, line),

        // ── .NET StreamReader / StreamWriter adapters — text I/O ────
        // Load-whole-file model: `new StreamReader(path)` materializes a
        // string buffer via `node:fs.readFileSync`, `new StreamWriter`
        // accumulates into `__buf` and flushes via `writeFileSync`.
        // Bytecode-only — no `dotnet:io` host fns.
        "dotnet.stream_reader_new"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_new(chunks, current, argc, line),
        "dotnet.stream_reader_read_line"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_read_line(chunks, current, line),
        "dotnet.stream_reader_read_to_end"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_read_to_end(chunks, current, line),
        "dotnet.stream_reader_at_end"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_at_end(chunks, current, line),
        "dotnet.stream_reader_close"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_reader_close(chunks, current, line),
        "dotnet.stream_writer_new"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_writer_new(chunks, current, argc, line),
        "dotnet.stream_writer_write"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_writer_write(chunks, current, line),
        "dotnet.stream_writer_write_line"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_writer_write_line(chunks, current, line),
        "dotnet.stream_writer_flush"
            => crate::emitter::dotnet::core::stream_io_adapter::emit_stream_writer_flush(chunks, current, line),

        // ── PHP `isset(...)` — variadic null check, returns true iff
        // ALL args are non-null. Inline emit folds an AND chain.
        "php.isset_all" => emit_isset_all(&mut chunks[current], argc, line),

        // ── Ruby `obj.dig(k1, k2, ..., kN)` — variadic property walk.
        // Returns `obj[k1]?[k2]?...[kN]`, or `nil` if any link is null.
        // `argc` includes receiver: `argc == N + 1`. Inline emit chains
        // ARRAY_GET (polymorphic over Map / Object / Array) with
        // null-short-circuit at every step.
        "ruby.dig" => emit_dig(&mut chunks[current], argc, line),

        // ── VB Choose / Switch — variadic 1-indexed selector ────────
        // `Choose(idx, v1, v2, ..., vN)` returns `vidx`. Variadic so it
        // needs `argc` threading; can't be a stdlib chunk (fixed arity).
        // Implementation: pack trailing vals into an array via
        // `ARRAY_NEW_FIXED`, save to a local, then `ARRAY_GET array[idx-1]`.
        // .NET-shape rather than stdlib because Choose is a VB.NET / VBA
        // language built-in, not a generic helper.
        "dotnet.choose" => emit_choose(&mut chunks[current], argc, line),
        "threading.thread_join" => threading::emit_thread_join(&mut chunks[current], line),
        "threading.atomic_load" => threading::emit_atomic_load(&mut chunks[current], line),
        "threading.atomic_store" => threading::emit_atomic_store(&mut chunks[current], line),
        "threading.atomic_add" => threading::emit_atomic_add(&mut chunks[current], line),
        "threading.atomic_sub" => threading::emit_atomic_sub(&mut chunks[current], line),
        "threading.atomic_xchg" => threading::emit_atomic_xchg(&mut chunks[current], line),
        "threading.atomic_cmpxchg" => threading::emit_atomic_cmpxchg(&mut chunks[current], line),
        "threading.atomic_fence" => threading::emit_atomic_fence(&mut chunks[current], line),
        "threading.suspend" => threading::emit_suspend(&mut chunks[current], line),

        _ => return false,
    }
    true
}

/// Handle common ops that need to register a host import in addition to
/// emitting bytecode. `import` is a callback that resolves an import to its
/// index (typically by adding to chunk[0]'s import table).
///
/// Returns `true` if `name` was recognized and emitted; on `true`, the
/// stack discipline matches the helper's contract (e.g. `threading.sleep`
/// leaves a `null` on the stack so the call site can drop it uniformly).
/// Returns `false` if the name is unknown OR doesn't need imports — call
/// `emit_common` for those.
pub fn emit_common_with_imports(
    name: &str,
    chunk: &mut Chunk,
    argc: u8,
    line: u32,
    mut import: impl FnMut(&str, &str) -> u16,
) -> bool {
    let _ = argc; // unused by current emits — kept for parity with `emit_common`
    match name {
        "threading.sleep" => {
            // Sleep uses the standard WASI clocks import. This is intentional:
            // wasi:clocks is the WASI standard for time/sleep and works on any
            // WASI-compliant runtime. Routed through the dispatcher so the
            // primitive can be swapped later (e.g. memory_atomic_wait32 with
            // a timeout, once shared-memory layout is settled) without
            // touching every language profile.
            let idx = import("wasi:clocks", "sleep");
            threading::emit_sleep(chunk, idx, line);
            // emit_sleep already drops the import return; push null so the
            // call site has a stack value to consume.
            chunk.emit_op(Op::NULL, line);
        }
        _ => return false,
    }
    true
}


/// Emit `Choose(idx, v1, v2, ..., vN)` — VB-style 1-indexed variadic
/// selector. Stack on entry: `[idx, v1, v2, ..., vN]` where `argc == N+1`.
/// Stack on exit: `[v_idx]` (the value at position `idx`, 1-indexed).
fn emit_choose(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc < 2 {
        // Defensive: with no values, just push null.
        chunk.emit_op(Op::NULL, line);
        return;
    }
    let n = (argc as u16) - 1;
    let arr_slot = chunk.local_count;
    let idx_slot = arr_slot + 1;
    chunk.local_count = arr_slot + 2;

    // Pack the top N values into an array. Stack: [idx, arr]
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, n, line);
    // Save array; stack: [idx]
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunk.emit_op(Op::DROP, line);
    // Convert idx (f64) to i32 then subtract 1 (Choose is 1-indexed).
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_op(Op::I32_CONST_1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunk.emit_op(Op::DROP, line);
    // Push [array, idx-1] for ARRAY_GET.
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// Emit PHP `isset($a, $b, $c)` — returns true iff every arg is non-null.
/// Stack on entry: `[v1, v2, ..., vN]` where `argc == N`. Stack on exit:
/// `[bool]`. Implementation folds `REF_IS_NULL → DYN_NOT` over each arg and
/// AND-chains the results — pure WASM, no host fns.
fn emit_isset_all(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc == 0 {
        // PHP `isset()` with no args → true (vacuous truth).
        chunk.emit_op(Op::TRUE, line);
        return;
    }
    // Stash all args in temps so we can fold them deterministically.
    let base = chunk.local_count;
    chunk.local_count = base + argc as u16;
    for i in (0..argc).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
        chunk.emit_op(Op::DROP, line);
    }
    // result = !is_null(args[0])
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::DYN_NOT, line);
    // result = result AND !is_null(args[i]) for each remaining
    for i in 1..argc {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::DYN_NOT, line);
        chunk.emit_op(Op::I32_AND, line);
    }
}

/// Emit Ruby `obj.dig(k1, k2, ..., kN)` — variadic property walk with
/// null short-circuit. Stack: `[receiver, k1, k2, ..., kN]` where
/// `argc == N + 1` (receiver + N keys). Stack on exit: `[value_or_null]`.
///
/// Strategy: stash all keys + receiver into temps. Walk one key at a
/// time using `Op::ARRAY_GET` (polymorphic Map/Array/Object). Between
/// each step, check if current value is null and short-circuit out of
/// the wrapping block if so.
fn emit_dig(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc == 0 {
        chunk.emit_op(Op::NULL, line);
        return;
    }
    if argc == 1 {
        // Just the receiver, no keys — return it as-is.
        return;
    }
    let nkeys = argc - 1;
    // Allocate temps: `cur` slot + N key slots.
    let cur_slot = chunk.local_count;
    chunk.local_count = cur_slot + 1 + nkeys as u16;
    // Stash keys back-to-front (last key first, ends up in highest slot).
    for i in (0..nkeys).rev() {
        let slot = cur_slot + 1 + i as u16;
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        chunk.emit_op(Op::DROP, line);
    }
    // Stash receiver as initial `cur`.
    chunk.emit_op_u16(Op::LOCAL_SET, cur_slot, line);
    chunk.emit_op(Op::DROP, line);

    // Wrapping block: `br_if(0)` exits early when `cur` becomes null.
    let exit_block = chunk.emit_block(line);
    for i in 0..nkeys {
        let key_slot = cur_slot + 1 + i as u16;
        // if cur is null: exit
        chunk.emit_op_u16(Op::LOCAL_GET, cur_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_br_if(0, line);
        // cur = ARRAY_GET(cur, key)
        chunk.emit_op_u16(Op::LOCAL_GET, cur_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, cur_slot, line);
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_end(line); chunk.patch_block(exit_block);

    // Push final result.
    chunk.emit_op_u16(Op::LOCAL_GET, cur_slot, line);
}
