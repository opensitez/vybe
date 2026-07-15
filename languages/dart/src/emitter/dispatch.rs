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
