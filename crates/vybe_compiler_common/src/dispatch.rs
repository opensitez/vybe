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

use crate::{collections, dict, strings, threading};

/// Handle common ops that need only a chunk and line.
/// Returns `true` if `name` was recognized and emitted, `false` otherwise.
pub fn emit_common(name: &str, chunk: &mut Chunk, line: u32) -> bool {
    match name {
        // ── Dict ops ──
        "dict.set_dynamic" => {
            dict::emit_set_dynamic(chunk, line);
            chunk.emit_op(Op::null, line); // void return
        }
        "dict.get_dynamic" => dict::emit_get_dynamic(chunk, line),
        "dict.has" => dict::emit_method_has(chunk, line),
        "dict.delete" => dict::emit_method_delete(chunk, line),
        "dict.clear" => dict::emit_method_clear_stack(chunk, line),
        "dict.size" => dict::emit_method_size(chunk, line),
        "dict.keys" => dict::emit_keys(chunk, line),
        "dict.values" => dict::emit_values(chunk, line),
        "dict.new" => dict::emit_new(chunk, line),

        // ── Collection ops ──
        "collections.push" => collections::emit_push(chunk, line),
        "collections.pop" => collections::emit_pop(chunk, line),
        "collections.length" => collections::emit_len(chunk, line),
        "collections.get" => collections::emit_get(chunk, line),
        "collections.set" => collections::emit_set(chunk, line),
        "collections.contains" => collections::emit_contains(chunk, line),
        "collections.index_of" => collections::emit_index_of(chunk, line),
        "collections.sorted" => collections::emit_sorted(chunk, line),
        "collections.reverse" => collections::emit_reverse(chunk, line),
        "collections.join" => collections::emit_join(chunk, line),
        "collections.slice" => collections::emit_slice(chunk, line),
        "collections.new" => collections::emit_array_new(chunk, 0, line),

        // ── String ops ──
        "strings.length" => strings::emit_length(chunk, line),
        "strings.to_upper" => strings::emit_to_upper(chunk, line),
        "strings.to_lower" => strings::emit_to_lower(chunk, line),
        "strings.trim" => strings::emit_trim(chunk, line),
        "strings.substring" => strings::emit_substring(chunk, line),
        "strings.replace" => strings::emit_replace(chunk, line),
        "strings.split" => strings::emit_split(chunk, line),
        "strings.index_of" => strings::emit_index_of(chunk, line),
        "strings.concat" => strings::emit_concat(chunk, 2, line),

        // ── Expression ops ──
        "expressions.undefined" => crate::expressions::emit_undefined(chunk, line),

        // ── Threading ops ──
        // Real WASM threading opcodes (wasi-threads proposal):
        // thread_spawn, thread_join, memory atomic_*. NOT host calls — these
        // run unchanged on any standard WASM runtime that supports the
        // threads proposal.
        "threading.task_run" => threading::emit_task_run(chunk, line),
        "threading.thread_new" => threading::emit_thread_new(chunk, line),
        "threading.thread_spawn" => threading::emit_thread_spawn(chunk, line),
        "threading.thread_join" => threading::emit_thread_join(chunk, line),
        "threading.atomic_load" => threading::emit_atomic_load(chunk, line),
        "threading.atomic_store" => threading::emit_atomic_store(chunk, line),
        "threading.atomic_add" => threading::emit_atomic_add(chunk, line),
        "threading.atomic_sub" => threading::emit_atomic_sub(chunk, line),
        "threading.atomic_xchg" => threading::emit_atomic_xchg(chunk, line),
        "threading.atomic_cmpxchg" => threading::emit_atomic_cmpxchg(chunk, line),
        "threading.atomic_fence" => threading::emit_atomic_fence(chunk, line),
        "threading.suspend" => threading::emit_suspend(chunk, line),

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
    line: u32,
    mut import: impl FnMut(&str, &str) -> u16,
) -> bool {
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
            chunk.emit_op(Op::null, line);
        }
        _ => return false,
    }
    true
}
