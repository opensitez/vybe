//! Pascal-specific common dispatch.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
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
