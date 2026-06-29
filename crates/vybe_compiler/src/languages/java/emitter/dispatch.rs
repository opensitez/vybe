//! Java-specific `common:java.*` dispatch.
//!
//! All patterns follow the same conventions as Go/Dart runtime adapters:
//! - Imports register on `chunks[0]`; code emits to `chunks[current]`
//! - `host::emit(chunk, module, name, argc, line)` for inline host calls
//! - `collections::emit_*(chunks, current, line)` for collection helpers
//! - `strings::emit_*(chunk, line)` for string helpers (single chunk)
//! - `core_wasm::*(&mut chunk, line, ...)` for raw WASM ops

use crate::emitter::instructions::{core_wasm, host};
use crate::emitter::{collections, strings};
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        // ── Print ──────────────────────────────────────────────────────────
        "java.print_no_newline" => {
            let to_str = chunks[0].add_import("ecma:string", "String");
            chunks[current].emit_op_u16(Op::CALL_IMPORT, to_str, line);
            chunks[current].emit(1, line);
            host::emit(&mut chunks[current], "wasi:cli", "print", 1, line);
        }
        "java.printf" => {
            crate::emitter::sprintf::emit_sprintf(chunks, current, argc, line);
            host::emit(&mut chunks[current], "wasi:logging/logging", "log", 1, line);
        }

        // ── String helpers ─────────────────────────────────────────────────
        "java.str_is_empty" => {
            // Polymorphic: works on String, List, Map.
            core_wasm::dup(&mut chunks[current], line);
            host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
            crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            strings::emit_length(&mut chunks[current], line);
            chunks[current].emit_else(line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_end(line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
        }
        "java.str_is_blank" => {
            strings::emit_trim(&mut chunks[current], line);
            strings::emit_length(&mut chunks[current], line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
        }
        "java.is_empty" => {
            collections::emit_len(chunks, current, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
        }
        "java.size" => {
            collections::emit_len(chunks, current, line);
        }
        "java.replace_regex" => {
            host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
        }
        "java.replace_first_regex" => {
            host::emit(&mut chunks[current], "ecma:string", "replace", 3, line);
        }
        "java.compare_ignore_case" => {
            host::emit(
                &mut chunks[current],
                "ecma:string",
                "compareIgnoreCase",
                2,
                line,
            );
        }
        "java.equals_ignore_case" => {
            host::emit(
                &mut chunks[current],
                "ecma:string",
                "equalsIgnoreCase",
                2,
                line,
            );
        }
        "java.str_matches" => {
            host::emit(&mut chunks[current], "ecma:string", "matches", 2, line);
        }
        "java.to_char_array" => {
            strings::emit_split(&mut chunks[current], line);
        }
        "java.str_lines" => {
            host::emit(&mut chunks[current], "ecma:string", "lines", 1, line);
        }

        // ── Numeric conversions ───────────────────────────────────────────
        "java.to_binary_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toBinary", 1, line);
        }
        "java.to_hex_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toHex", 1, line);
        }
        "java.to_octal_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toOctal", 1, line);
        }
        "java.is_infinite" => {
            host::emit(&mut chunks[current], "ecma:number", "isFinite", 1, line);
            crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
        }
        "java.signum" => {
            host::emit(&mut chunks[current], "ecma:math", "sign", 1, line);
        }
        "java.floor_div" => {
            host::emit(&mut chunks[current], "ecma:math", "floorDiv", 2, line);
        }
        "java.floor_mod" => {
            host::emit(&mut chunks[current], "ecma:math", "floorMod", 2, line);
        }

        // ── Character helpers ─────────────────────────────────────────────
        "java.char_is_digit" => {
            host::emit(&mut chunks[current], "ecma:char", "isDigit", 1, line);
        }
        "java.char_is_letter" => {
            host::emit(&mut chunks[current], "ecma:char", "isLetter", 1, line);
        }
        "java.char_is_alnum" => {
            host::emit(&mut chunks[current], "ecma:char", "isAlnum", 1, line);
        }
        "java.char_is_upper" => {
            host::emit(&mut chunks[current], "ecma:char", "isUpper", 1, line);
        }
        "java.char_is_lower" => {
            host::emit(&mut chunks[current], "ecma:char", "isLower", 1, line);
        }
        "java.char_is_space" => {
            host::emit(&mut chunks[current], "ecma:char", "isSpace", 1, line);
        }
        "java.char_to_upper" => {
            strings::emit_to_upper(&mut chunks[current], line);
        }
        "java.char_to_lower" => {
            strings::emit_to_lower(&mut chunks[current], line);
        }
        "java.char_numeric" => {
            host::emit(&mut chunks[current], "ecma:number", "parseInt", 1, line);
        }

        // ── Array helpers ─────────────────────────────────────────────────
        "java.new_array" => {
            collections::emit_new_with_length(chunks, current, line);
        }
        "java.array_clone" => {
            collections::emit_slice(chunks, current, line);
        }
        "java.arrays_sort" => {
            collections::emit_sort(chunks, current, line);
        }
        "java.arrays_fill" => {
            collections::emit_fill(chunks, current, line);
        }
        "java.arrays_copy_of" => {
            collections::emit_slice(chunks, current, line);
        }
        "java.arrays_copy_of_range" => {
            collections::emit_get_range(chunks, current, line);
        }
        "java.arrays_to_string" => {
            host::emit(&mut chunks[current], "ecma:array", "toString", 1, line);
        }
        "java.arrays_equals" => {
            host::emit(&mut chunks[current], "ecma:array", "equals", 2, line);
        }
        "java.arrays_binary_search" => {
            collections::emit_index_of(chunks, current, line);
        }

        // ── List helpers ──────────────────────────────────────────────────
        "java.list_of" | "java.set_of" => { /* varargs arrive as array — noop */ }
        "java.map_of" => {
            host::emit(&mut chunks[current], "ecma:object", "fromPairs", argc, line);
        }
        "java.map_entry" => {
            host::emit(&mut chunks[current], "ecma:array", "pair", 2, line);
        }
        "java.empty_list" => {
            collections::emit_array_new(chunks, current, 0, line);
        }
        "java.singleton_list" => {
            host::emit(&mut chunks[current], "ecma:array", "of", 1, line);
        }
        "java.n_copies" => {
            host::emit(&mut chunks[current], "ecma:array", "nCopies", 2, line);
        }
        "java.list_get" => {
            collections::emit_get(chunks, current, line);
        }
        "java.list_set" => {
            collections::emit_set(chunks, current, line);
        }
        "java.list_remove" => {
            collections::emit_remove_at(chunks, current, line);
        }
        "java.list_clear" => {
            host::emit(&mut chunks[current], "ecma:array", "clear", 1, line);
        }
        "java.list_contains" => {
            collections::emit_contains(chunks, current, line);
        }
        "java.sub_list" => {
            collections::emit_slice(chunks, current, line);
        }
        "java.list_sort" => {
            collections::emit_sort(chunks, current, line);
        }
        "java.add_all" => {
            collections::emit_concat(chunks, current, line);
        }
        "java.remove_all" => {
            host::emit(&mut chunks[current], "ecma:array", "removeAll", 2, line);
        }
        "java.retain_all" => {
            host::emit(&mut chunks[current], "ecma:array", "retainAll", 2, line);
        }
        "java.list_for_each" => {
            // Delegate to runtime helper; avoids needing pre-allocated local slots.
            collections::emit_runtime_helper_call(chunks, current, "__array_for_each", argc, line);
        }
        "java.queue_poll" => {
            collections::emit_shift(chunks, current, line);
        }
        "java.add_first" => {
            host::emit(&mut chunks[current], "ecma:array", "unshift", 2, line);
        }
        "java.remove_first" => {
            collections::emit_shift(chunks, current, line);
        }
        "java.peek_first" => {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            collections::emit_get(chunks, current, line);
        }
        "java.peek_last" => {
            core_wasm::dup(&mut chunks[current], line);
            collections::emit_len(chunks, current, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_SUB, line);
            collections::emit_get(chunks, current, line);
        }

        // ── Map helpers ────────────────────────────────────────────────────
        "java.map_put" => {
            crate::emitter::dict::emit_set(chunks, current, line);
        }
        "java.map_put_all" => {
            host::emit(&mut chunks[current], "ecma:object", "assign", 2, line);
        }
        "java.map_get_or_default" => {
            host::emit(&mut chunks[current], "ecma:object", "getOrDefault", 3, line);
        }
        "java.map_contains_key" => {
            crate::emitter::dict::emit_method_has(chunks, current, line);
        }
        "java.map_contains_value" => {
            host::emit(&mut chunks[current], "ecma:object", "hasValue", 2, line);
        }
        "java.entry_set" => {
            collections::emit_iter_entries(chunks, current, line);
        }
        "java.put_if_absent" => {
            host::emit(&mut chunks[current], "ecma:object", "putIfAbsent", 3, line);
        }
        "java.compute_if_absent" => {
            host::emit(
                &mut chunks[current],
                "ecma:object",
                "computeIfAbsent",
                3,
                line,
            );
        }
        "java.map_merge" => {
            host::emit(&mut chunks[current], "ecma:object", "merge", 4, line);
        }
        "java.map_remove" => {
            crate::emitter::dict::emit_method_delete(chunks, current, line);
        }
        "java.map_replace" => {
            crate::emitter::dict::emit_set(chunks, current, line);
        }
        "java.entry_key" => {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            collections::emit_get(chunks, current, line);
        }
        "java.entry_value" => {
            core_wasm::i32_const(&mut chunks[current], line, 1);
            collections::emit_get(chunks, current, line);
        }

        // ── StringBuilder helpers ──────────────────────────────────────────
        "java.sb_append" => {
            crate::emitter::dotnet::dispatch::dispatch(
                "dotnet.sb_append",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_insert" => {
            crate::emitter::dotnet::dispatch::dispatch(
                "dotnet.sb_insert",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_delete" => {
            crate::emitter::dotnet::dispatch::dispatch(
                "dotnet.sb_remove",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.sb_delete_char_at" => {
            crate::emitter::dotnet::dispatch::dispatch(
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
        "java.collections_sort" => {
            collections::emit_sort(chunks, current, line);
        }
        "java.collections_reverse" => {
            collections::emit_reverse(chunks, current, line);
        }
        "java.collections_shuffle" => {
            host::emit(&mut chunks[current], "ecma:array", "shuffle", 1, line);
        }
        "java.collections_min" => {
            host::emit(&mut chunks[current], "ecma:array", "min", 1, line);
        }
        "java.collections_max" => {
            host::emit(&mut chunks[current], "ecma:array", "max", 1, line);
        }
        "java.collections_frequency" => {
            host::emit(&mut chunks[current], "ecma:array", "frequency", 2, line);
        }

        // ── String formatting ─────────────────────────────────────────────
        "java.string_format" => {
            crate::emitter::dotnet::dispatch::dispatch(
                "dotnet.string_format",
                chunks,
                current,
                argc,
                line,
            );
        }
        "java.string_join" => {
            collections::emit_join(chunks, current, line);
        }

        // ── Optional ─────────────────────────────────────────────────────
        "java.optional_or_else" | "java.optional_or_else_get" => {
            host::emit(&mut chunks[current], "ecma:optional", "orElse", 2, line);
        }
        "java.optional_is_present" => {
            host::emit(&mut chunks[current], "ecma:optional", "isPresent", 1, line);
        }
        "java.optional_if_present" => {
            host::emit(&mut chunks[current], "ecma:optional", "ifPresent", 2, line);
        }
        "java.optional_or_else_throw" => {
            host::emit(
                &mut chunks[current],
                "ecma:optional",
                "orElseThrow",
                1,
                line,
            );
        }

        // ── Object utilities ──────────────────────────────────────────────
        "java.equals" => {
            crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
        }
        "java.hash_code" => {
            host::emit(&mut chunks[current], "ecma:object", "hashCode", 1, line);
        }
        "java.require_non_null" => {
            host::emit(
                &mut chunks[current],
                "ecma:object",
                "requireNonNull",
                argc,
                line,
            );
        }
        "java.is_null" => {
            host::emit(&mut chunks[current], "ecma:object", "isNull", 1, line);
        }
        "java.non_null" => {
            host::emit(&mut chunks[current], "ecma:object", "nonNull", 1, line);
        }

        _ => return false,
    }
    true
}
