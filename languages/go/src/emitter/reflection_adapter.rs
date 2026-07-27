//! Go `reflect` facade routed through the shared reflection substrate.

use vybe_bytecode::Chunk;
use vybe_compiler::compiler::reflection;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.reflect_typeof" => {
            reflection::emit_type_descriptor_from_stack(chunks, current, argc, line)
        }
        "go.reflect_valueof" => {
            reflection::emit_value_descriptor_from_stack(chunks, current, argc, line)
        }
        "go.reflect_kind" => {
            reflection::emit_descriptor_field(&mut chunks[current], reflection::FIELD_KIND, line)
        }
        "go.reflect_name" => reflection::emit_descriptor_field(
            &mut chunks[current],
            reflection::FIELD_TYPE_NAME,
            line,
        ),
        "go.reflect_interface" => {
            reflection::emit_descriptor_field(&mut chunks[current], reflection::FIELD_VALUE, line)
        }
        "go.reflect_int" | "go.reflect_uint" | "go.reflect_float" | "go.reflect_bool"
        | "go.reflect_string" => {
            reflection::emit_descriptor_field(&mut chunks[current], reflection::FIELD_VALUE, line)
        }
        "go.reflect_num_field" => reflection::emit_reflect_num_field(chunks, current, line),
        "go.reflect_field" => reflection::emit_reflect_field(chunks, current, line),
        "go.reflect_field_by_name" => reflection::emit_reflect_field_by_name(chunks, current, line),
        "go.reflect_len" => reflection::emit_reflect_len(chunks, current, line),
        "go.reflect_index" => reflection::emit_reflect_index(chunks, current, line),
        "go.reflect_map_index" => reflection::emit_reflect_map_index(chunks, current, line),
        "go.reflect_is_valid" => reflection::emit_reflect_is_valid(chunks, current, line),
        "go.reflect_is_nil" => reflection::emit_reflect_is_nil(chunks, current, line),
        "go.reflect_can_set" => reflection::emit_reflect_can_set(chunks, current, line),
        "go.reflect_is_zero" => reflection::emit_reflect_is_zero(chunks, current, line),
        "go.reflect_elem" => reflection::emit_reflect_elem(chunks, current, line),
        "go.reflect_set" => reflection::emit_reflect_set_value(chunks, current, line),
        "go.reflect_set_int"
        | "go.reflect_set_uint"
        | "go.reflect_set_string"
        | "go.reflect_set_bool" => reflection::emit_reflect_set_primitive(chunks, current, line),
        _ => return false,
    }
    true
}
