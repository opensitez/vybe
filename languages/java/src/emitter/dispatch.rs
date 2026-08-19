//! Java-specific `common:java.*` dispatch.
//!
//! All patterns follow the same conventions as Go/Dart runtime adapters:
//! - Imports register on `chunks[0]`; code emits to `chunks[current]`
//! - `host::emit(chunk, module, name, argc, line)` for inline host calls
//! - `collections::emit_*(chunks, current, line)` for collection helpers
//! - `strings::emit_*(chunk, line)` for string helpers (single chunk)
//! - `core_wasm::*(&mut chunk, line, ...)` for raw WASM ops

use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::{collections, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn emit_stdout_text(chunk: &mut Chunk, line: u32) {
    let text_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    host::emit(chunk, "wasi:cli/stdout", "get-stdout", 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    host::emit(
        chunk,
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
        2,
        line,
    );
}

fn emit_print_stream_sentinel(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("__java_out", line);
}

fn emit_java_exp(chunks: &mut [Chunk], current: usize, upper: bool, line: u32) {
    chunks[current].emit_i32_const(6, line);
    let to_exp = chunks[current].add_import("ecma:number", "toExponential");
    chunks[current].emit_call(to_exp, 2, line);
    if upper {
        let to_upper = chunks[current].add_import("ecma:string", "toUpperCase");
        chunks[current].emit_call(to_upper, 1, line);
    }

    let (plus, plus_padded, minus, minus_padded) = if upper {
        ("E+", "E+0", "E-", "E-0")
    } else {
        ("e+", "e+0", "e-", "e-0")
    };
    chunks[current].emit_string_const(plus, line);
    chunks[current].emit_string_const(plus_padded, line);
    strings::emit_replace(&mut chunks[current], line);
    chunks[current].emit_string_const(minus, line);
    chunks[current].emit_string_const(minus_padded, line);
    strings::emit_replace(&mut chunks[current], line);
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "java.println" => {
            if argc == 0 {
                chunks[current].emit_string_const("", line);
            } else {
                // `println(Object)` IS `String.valueOf(x)` (java.io.PrintStream),
                // which reaches the object's ToString slot.
                vybe_platform_jvm::emitter::string_adapter::emit_value_of(chunks, current, line);
            }
            host::emit(&mut chunks[current], "web:console", "log", 1, line);
            emit_print_stream_sentinel(&mut chunks[current], line);
        }
        "java.print_no_newline" => {
            if argc == 0 {
                chunks[current].emit_string_const("", line);
            }
            // Real WASI stdout (the old target, `wasi:cli.print`, never
            // existed as a host fn).
            vybe_platform_jvm::emitter::string_adapter::emit_value_of(chunks, current, line);
            emit_stdout_text(&mut chunks[current], line);
            emit_print_stream_sentinel(&mut chunks[current], line);
        }
        "java.field_set" => {
            let value = chunks[current].alloc_scratch(1);
            let field = chunks[current].alloc_scratch(1);
            let object = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
            host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
        }
        "java.field_inc" => {
            let delta = chunks[current].alloc_scratch(1);
            let field = chunks[current].alloc_scratch(1);
            let object = chunks[current].alloc_scratch(1);
            let value = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delta, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
            host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, delta, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
            host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
        }
        "java.printf" => {
            vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
            emit_stdout_text(&mut chunks[current], line);
            emit_print_stream_sentinel(&mut chunks[current], line);
        }
        "java.printf_array" => {
            vybe_compiler::primitives::sprintf::emit_sprintf_from_array(chunks, current, line);
            emit_stdout_text(&mut chunks[current], line);
            emit_print_stream_sentinel(&mut chunks[current], line);
        }
        "java.format_grouped_int" => {
            let to_locale = chunks[current].add_import("ecma:number", "toLocaleString");
            chunks[current].emit_call(to_locale, 1, line);
        }
        "java.format_exp_lower" => {
            emit_java_exp(chunks, current, false, line);
        }
        "java.format_exp_upper" => {
            emit_java_exp(chunks, current, true, line);
        }
        "java.str_is_empty" => {
            // Polymorphic: works on String, List, Map.
            core_wasm::dup(&mut chunks[current], line);
            host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            strings::emit_length(&mut chunks[current], line);
            chunks[current].emit_else(line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_end(line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "java.str_is_blank" => {
            strings::emit_trim(&mut chunks[current], line);
            strings::emit_length(&mut chunks[current], line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "java.new_array" => {
            collections::emit_new_with_length(chunks, current, line);
        }
        "java.new_int_array" => {
            emit_new_array_with_default(chunks, current, line, JavaArrayDefault::IntZero);
        }
        "java.new_bool_array" => {
            emit_new_array_with_default(chunks, current, line, JavaArrayDefault::BoolFalse);
        }
        "java.new_int_2d_array" => {
            vybe_platform_jvm::emitter::arrays_adapter::emit_new_int_2d(chunks, current, line);
        }
        "java.array_clone" => {
            collections::emit_slice(chunks, current, line);
        }
        "java.mutable_list_of" => {
            collections::emit_array_new(chunks, current, argc as u16, line);
        }
        "java.list_contains" => {
            collections::emit_contains(chunks, current, line);
        }
        "java.entry_key" => {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            collections::emit_get(chunks, current, line);
        }
        "java.entry_value" => {
            core_wasm::i32_const(&mut chunks[current], line, 1);
            collections::emit_get(chunks, current, line);
        }
        "java.sb_delete_char_at" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.sb_remove",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.equals" => {
            vybe_compiler::primitives::object::emit_equals(&mut chunks[current], line);
        }
        "java.hash_code" => {
            vybe_compiler::primitives::object::emit_hash_code(&mut chunks[current], line);
        }
        // ── Print ──────────────────────────────────────────────────────────

        // ── Random helpers ────────────────────────────────────────────────

        // ── String helpers ─────────────────────────────────────────────────
        "java.is_empty" => {
            collections::emit_len(chunks, current, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "java.size" => {
            vybe_platform_jvm::emitter::list_adapter::emit_size(chunks, current, line);
        }

        // ── Numeric conversions ───────────────────────────────────────────

        // ── Integer bit operations (JLS java.lang.Integer) — raw WASM ops ──
        "java.compare" => {
            let b_slot = chunks[current].alloc_scratch(1);
            let a_slot = chunks[current].alloc_scratch(1);
            let result_slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
            vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(-1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
            vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
        }

        // ── Character helpers ─────────────────────────────────────────────
        // `ecma:char` was never a registered host module — these arms
        // panicked at compile time. Shared with kotlin via the jvm
        // platform's guards over the tier-3 `strings::emit_is_*` primitives.

        // ── Array helpers ─────────────────────────────────────────────────

        // ── Primitive stream helpers ─────────────────────────────────────

        // ── List helpers ──────────────────────────────────────────────────
        "java.hash_set_new" => {
            vybe_platform_jvm::emitter::collection_adapter::emit_hash_set_new(
                chunks, current, argc, line,
            );
        }
        "java.add" => {
            vybe_platform_jvm::emitter::list_adapter::emit_add(chunks, current, argc, line);
        }
        "java.map_get" => {
            vybe_platform_jvm::emitter::list_adapter::emit_map_get(chunks, current, line);
        }
        "java.concurrent_hash_map_new" => {
            vybe_platform_jvm::emitter::list_adapter::emit_concurrent_hash_map_new(chunks, current, argc, line);
        }
        "java.get" => {
            if argc <= 1 {
                vybe_platform_jvm::emitter::stream_adapter::emit_get_optional_value(
                    chunks, current, line,
                );
            } else {
                vybe_platform_jvm::emitter::list_adapter::emit_get_or_map_get(chunks, current, line);
            }
        }
        "java.list_set" => {
            vybe_platform_jvm::emitter::list_adapter::emit_set(chunks, current, argc, line);
        }
        "java.list_remove" => {
            vybe_platform_jvm::emitter::list_adapter::emit_remove_at(chunks, current, line);
        }
        "java.list_clear" => {
            vybe_platform_jvm::emitter::list_adapter::emit_clear(chunks, current, line);
        }
        "java.sorted_first" => {
            vybe_platform_jvm::emitter::list_adapter::emit_sorted_end(chunks, current, false, line);
        }
        "java.sorted_last" => {
            vybe_platform_jvm::emitter::list_adapter::emit_sorted_end(chunks, current, true, line);
        }
        "java.remove_first" => {
            collections::emit_shift(chunks, current, line);
        }
        "java.poll_last" => {
            vybe_platform_jvm::emitter::list_adapter::emit_poll(chunks, current, true, line);
        }

        // ── Map helpers ────────────────────────────────────────────────────
        "java.map_put" => {
            vybe_platform_jvm::emitter::list_adapter::emit_map_put(chunks, current, line);
        }
        "java.map_key_set" => {
            vybe_platform_jvm::emitter::list_adapter::emit_map_key_set(chunks, current, line);
        }
        "java.entry_set" => {
            vybe_platform_jvm::emitter::list_adapter::emit_map_entry_set(chunks, current, line);
        }
        "java.list_iterator" => {
            vybe_platform_jvm::emitter::list_adapter::emit_list_iterator(chunks, current, argc, line);
        }
        "java.iterator_next" => {
            vybe_platform_jvm::emitter::list_adapter::emit_iterator_next(chunks, current, line);
        }
        "java.iterator_has_next" => {
            vybe_platform_jvm::emitter::list_adapter::emit_iterator_has_next(chunks, current, line);
        }
        "java.iterator_previous" => {
            vybe_platform_jvm::emitter::list_adapter::emit_iterator_previous(chunks, current, line);
        }
        "java.iterator_has_previous" => {
            vybe_platform_jvm::emitter::list_adapter::emit_iterator_has_previous(chunks, current, line);
        }
        "java.iterator_next_index" => {
            vybe_platform_jvm::emitter::list_adapter::emit_iterator_next_index(chunks, current, line);
        }
        "java.iterator_previous_index" => {
            vybe_platform_jvm::emitter::list_adapter::emit_iterator_previous_index(chunks, current, line);
        }

        // ── StringBuilder helpers ──────────────────────────────────────────
        "java.sb_append" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.sb_append",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_insert" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.sb_insert",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_delete" => {
            vybe_platform_dotnet::emitter::dispatch::dispatch(
                "dotnet.sb_remove",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_reverse" => {
            collections::emit_reverse(chunks, current, line);
        }

        // ── Collections utilities ─────────────────────────────────────────

        // ── String formatting ─────────────────────────────────────────────

        // ── Optional ─────────────────────────────────────────────────────

        // ── Object utilities ──────────────────────────────────────────────

        _ => return false,
    }
    true
}

enum JavaArrayDefault {
    IntZero,
    BoolFalse,
}

fn emit_new_array_with_default(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    default: JavaArrayDefault,
) {
    let len_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    collections::emit_new_with_length(chunks, current, line);

    let arr_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    match default {
        JavaArrayDefault::IntZero => chunks[current].emit_i32_const(0, line),
        JavaArrayDefault::BoolFalse => chunks[current].emit_bool_const(false, line),
    }
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    collections::emit_fill(chunks, current, line);
}
