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

use crate::emitter::{channels, collections, dict, strings, threading};

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
pub fn emit_common(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    // Language-specific routing lives in each language's emitter module,
    // registered the same way file extensions are: a language sets its
    // `emit_dispatch` in the `languages::all()` registry (and shared
    // platforms like `dotnet` register via `emitter::platform_emit_dispatch`).
    // The common dispatcher here only owns the genuinely-shared
    // `common:<cat>.*` keys below (collections/dict/strings/threading/…).
    // Adding a language never touches this file.
    if let Some(dot) = name.find('.') {
        if let Some(dispatch) = crate::languages::emit_dispatch_for(&name[..dot]) {
            if dispatch(name, chunks, current, argc, line) {
                return true;
            }
        }
    }
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
        "dict.items" => dict::emit_items(chunks, current, line),
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
        "collections.join_sep_first" => collections::emit_join_sep_first(chunks, current, line),
        "collections.slice" => collections::emit_slice(chunks, current, line),
        "collections.new" => collections::emit_array_new(chunks, current, 0, line),
        "collections.shift" => collections::emit_shift(chunks, current, line),
        "collections.concat" => collections::emit_concat(chunks, current, line),
        "collections.fill" => collections::emit_fill(chunks, current, line),
        "collections.sort" => collections::emit_sort(chunks, current, line),
        "collections.index_of_from" => collections::emit_index_of_from(chunks, current, line),
        "collections.last_index_of_from" => {
            collections::emit_last_index_of_from(chunks, current, line)
        }
        "collections.remove_range" => collections::emit_remove_range(chunks, current, line),
        "collections.get_range" => collections::emit_get_range(chunks, current, line),
        "collections.clone" => collections::emit_clone(chunks, current, line),
        "collections.sequence_equal" => collections::emit_sequence_equal(chunks, current, line),
        "collections.insert_range" => collections::emit_insert_range(chunks, current, line),
        "collections.set_range" => collections::emit_set_range(chunks, current, line),
        "collections.binary_search" => collections::emit_binary_search(chunks, current, line),
        "collections.reverse_range" => collections::emit_reverse_range(chunks, current, line),
        "collections.remove" => collections::emit_remove_value(chunks, current, line),
        "collections.insert" => collections::emit_insert_at(chunks, current, line),
        "collections.clear" => collections::emit_clear(chunks, current, line),
        "collections.sum" => collections::emit_sum(chunks, current, line),
        "collections.min" => collections::emit_pymin(chunks, current, line),
        "collections.max" => collections::emit_pymax(chunks, current, line),

        // ── Channel ops ──
        "channels.send" => channels::emit_send(chunks, current, line),
        "channels.receive" => channels::emit_receive(chunks, current, line),
        "channels.len" => channels::emit_len(chunks, current, line),
        "channels.cap" => channels::emit_cap(chunks, current, line),
        "channels.close" => channels::emit_close(chunks, current, line),

        // ── Python adapters ──
        "strings.length" => strings::emit_length(&mut chunks[current], line),
        "strings.to_upper" => strings::emit_to_upper(&mut chunks[current], line),
        "strings.to_lower" => strings::emit_to_lower(&mut chunks[current], line),
        "strings.trim" => strings::emit_trim(&mut chunks[current], line),
        "strings.substring" => strings::emit_substring(&mut chunks[current], line),
        "strings.replace" => strings::emit_replace(&mut chunks[current], line),
        "strings.split" => strings::emit_split(&mut chunks[current], line),
        "strings.index_of" => strings::emit_index_of(&mut chunks[current], line),
        "strings.concat" => strings::emit_concat(&mut chunks[current], 2, line),
        "sprintf.format" => crate::emitter::sprintf::emit_sprintf(chunks, current, argc, line),
        "sprintf.format_array" => {
            crate::emitter::sprintf::emit_sprintf_from_array(chunks, current, line)
        }

        // ── Expression ops ──
        "expressions.undefined" => {
            crate::emitter::expressions::emit_undefined(&mut chunks[current], line)
        }
        "expressions.i32_not" => {
            crate::emitter::expressions::emit_i32_not(&mut chunks[current], line)
        }
        "expressions.f64_mod" => {
            crate::emitter::expressions::emit_f64_mod(&mut chunks[current], line)
        }
        "expressions.bool_not" => {
            crate::emitter::expressions::emit_bool_not(&mut chunks[current], line)
        }

        // ── Delegate ops ──
        "delegates.combine" => crate::emitter::delegates::emit_combine(chunks, current, line),
        "delegates.remove" => crate::emitter::delegates::emit_remove(chunks, current, line),

        // ── JS Node compatibility adapters ───────────────────────
        // JS source keeps the Node-shaped call surface, but lowers
        // through compile-time adapters that compose the real
        // `wasi:sockets/*` interfaces. These live in the shared
        // `.NET` adapter home under `platforms/dotnet/core` so every
        // frontend can reuse them without a JS-only emitter fork.
        "threading.task_run" => threading::emit_task_run(chunks, current, line),
        "threading.task_delay" => threading::emit_task_delay(chunks, current, line),
        "threading.thread_new" => threading::emit_thread_new(chunks, current, line),
        "threading.thread_start" => threading::emit_thread_start(&mut chunks[current], line),
        "threading.thread_join" => threading::emit_thread_join(&mut chunks[current], line),
        "threading.thread_spawn" => threading::emit_thread_spawn(chunks, current, line),

        // ── VB Choose / Switch — variadic 1-indexed selector ────────
        // `Choose(idx, v1, v2, ..., vN)` returns `vidx`. Variadic so it
        // needs `argc` threading; can't be a stdlib chunk (fixed arity).
        // Implementation: pack trailing vals into an array via
        // `ARRAY_NEW_FIXED`, save to a local, then `ARRAY_GET array[idx-1]`.
        // .NET-shape rather than stdlib because Choose is a VB.NET / VBA
        // language built-in, not a generic helper.
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
