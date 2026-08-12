//! Pascal-specific common dispatch.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        // Delphi's `Generics.Collections` members that have no shared concept,
        // or that need an argument a `CommonEmit` name cannot carry. Each one
        // decomposes into the same `collections.*` routes — see
        // `runtime_adapter.rs`.
        "pascal.list_first" => {
            crate::emitter::runtime_adapter::emit_list_first(chunks, current, line);
            return true;
        }
        "pascal.list_last" => {
            crate::emitter::runtime_adapter::emit_list_last(chunks, current, line);
            return true;
        }
        "pascal.list_exchange" => {
            crate::emitter::runtime_adapter::emit_list_exchange(chunks, current, line);
            return true;
        }
        "pascal.list_move" => {
            crate::emitter::runtime_adapter::emit_list_move(chunks, current, line);
            return true;
        }
        "pascal.list_add_range" => {
            crate::emitter::runtime_adapter::emit_list_add_range(chunks, current, line);
            return true;
        }
        "pascal.list_extract_at" => {
            crate::emitter::runtime_adapter::emit_list_extract_at(chunks, current, line);
            return true;
        }
        "pascal.list_extract" => {
            crate::emitter::runtime_adapter::emit_list_extract(chunks, current, line);
            return true;
        }
        "pascal.list_noop" => {
            crate::emitter::runtime_adapter::emit_list_drop_args(chunks, current, argc, line);
            return true;
        }
        "pascal.self" => return true,
        // Every `E*` spelling routes to the SHARED exception constructor with
        // its canonical name bound at registration — see `exceptions.rs`.
        _ if name.starts_with("pascal.exc_") => {
            let key = &name["pascal.exc_".len()..];
            if let Some((spelling, _)) = crate::exceptions::EXCEPTION_TYPES
                .iter()
                .find(|(s, _)| s.to_lowercase() == key)
            {
                crate::emitter::runtime_adapter::emit_exception_new(
                    chunks, current, spelling, line,
                );
                return true;
            }
            return false;
        }
        // `TDictionary` over the shared Map — see `runtime_adapter.rs`.
        "pascal.dict_new" => {
            crate::emitter::runtime_adapter::emit_dict_new(chunks, current, line);
            return true;
        }
        "pascal.dict_size" => {
            crate::emitter::runtime_adapter::emit_dict_size(chunks, current, line);
            return true;
        }
        "pascal.dict_has" => {
            crate::emitter::runtime_adapter::emit_dict_has(chunks, current, line);
            return true;
        }
        "pascal.dict_contains_value" => {
            crate::emitter::runtime_adapter::emit_dict_contains_value(chunks, current, line);
            return true;
        }
        "pascal.dict_delete" => {
            crate::emitter::runtime_adapter::emit_dict_delete(chunks, current, line);
            return true;
        }
        "pascal.dict_clear" => {
            crate::emitter::runtime_adapter::emit_dict_clear(chunks, current, line);
            return true;
        }
        "pascal.dict_keys" => {
            crate::emitter::runtime_adapter::emit_dict_enumerate(chunks, current, "keys", line);
            return true;
        }
        "pascal.dict_values" => {
            crate::emitter::runtime_adapter::emit_dict_enumerate(chunks, current, "values", line);
            return true;
        }
        "pascal.dict_items" => {
            crate::emitter::runtime_adapter::emit_dict_enumerate(chunks, current, "entries", line);
            return true;
        }
        "pascal.str_delete" => {
            crate::emitter::runtime_adapter::emit_str_delete(chunks, current, line);
            return true;
        }
        "pascal.str_insert_var" => {
            crate::emitter::runtime_adapter::emit_str_insert_var(chunks, current, line);
            return true;
        }
        "pascal.trunc_cast" => {
            crate::emitter::runtime_adapter::emit_trunc_cast(chunks, current, line);
            return true;
        }
        "pascal.int_to_hex" => {
            crate::emitter::runtime_adapter::emit_int_to_hex(chunks, current, argc, line);
            return true;
        }
        "pascal.bool_to_str" => {
            crate::emitter::runtime_adapter::emit_bool_to_str(chunks, current, argc, line);
            return true;
        }
        "pascal.ansi_upper" => {
            crate::emitter::runtime_adapter::emit_ansi_case(chunks, current, true, line);
            return true;
        }
        "pascal.ansi_lower" => {
            crate::emitter::runtime_adapter::emit_ansi_case(chunks, current, false, line);
            return true;
        }
        "pascal.rgb" => {
            crate::emitter::runtime_adapter::emit_rgb(chunks, current, line);
            return true;
        }
        "pascal.extract_file_ext" => {
            crate::emitter::runtime_adapter::emit_extract_file_ext(chunks, current, line);
            return true;
        }
        "pascal.same_str" => {
            crate::emitter::runtime_adapter::emit_same_str(chunks, current, line);
            return true;
        }
        "pascal.same_text" => {
            crate::emitter::runtime_adapter::emit_same_text(chunks, current, true, line);
            return true;
        }
        "pascal.compare_text" => {
            crate::emitter::runtime_adapter::emit_same_text(chunks, current, false, line);
            return true;
        }
        "pascal.str_to_bool" => {
            crate::emitter::runtime_adapter::emit_str_to_bool(chunks, current, line);
            return true;
        }
        "pascal.str_to_int_def" => {
            crate::emitter::runtime_adapter::emit_str_to_int_def(chunks, current, line);
            return true;
        }
        "pascal.str_to_float_def" => {
            crate::emitter::runtime_adapter::emit_str_to_float_def(chunks, current, line);
            return true;
        }
        "pascal.write" => {
            crate::emitter::runtime_adapter::emit_pascal_write(chunks, current, argc, false, line);
            return true;
        }
        "pascal.writeln" => {
            crate::emitter::runtime_adapter::emit_pascal_write(chunks, current, argc, true, line);
            return true;
        }
        _ => {}
    }
    crate::emitter::runtime_adapter::emit_helper(name, chunks, current, argc, line)
}
