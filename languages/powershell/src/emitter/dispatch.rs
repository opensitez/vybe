//! `common:powershell.*` dispatch.
//!
//! Two arms, both proven. Every other PowerShell surface reaches a shared
//! primitive through the profile; these are the only places where none of them
//! can answer, which is the bar `documentation/powershellplan.md` §6.5 sets for
//! adding an emitter at all. Adding a third arm means proving it again.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "powershell.write_record" => {
            vybe_compiler::primitives::io::emit_print(&mut chunks[current], argc, line)
        }
        "powershell.add" => super::operators::emit_add(chunks, current, line),
        "powershell.collection_add" => {
            super::operators::emit_collection_add(chunks, current, argc, line)
        }
        "powershell.ensure_array" => {
            super::operators::emit_ensure_array(chunks, current, argc, line)
        }
        // Routes to `primitives::json`, which has no `common:json.*` dispatch
        // key — go and pascal reach it the same way.
        // `[builtin_slots.string] to_string` — the interpolation/concat slot.
        "powershell.get_enumerator" => super::operators::emit_get_enumerator(chunks, current, line),
        "powershell.unique_adjacent" => {
            super::operators::emit_unique_adjacent(chunks, current, line)
        }
        "powershell.divide" => super::operators::emit_divide(chunks, current, line),
        // `*` is the same LEFT-OPERAND rule `powershell.add` implements for
        // `+`, and needs it for the same reason: `F64_MUL` coerces an array or
        // a string operand to a number, so `@(1,2) * 3` answered 0 elements
        // and `"ab" * 3` answered `NaN`.
        "powershell.multiply" => super::operators::emit_multiply(chunks, current, line),
        "powershell.psobject" => super::operators::emit_psobject(chunks, current, line),
        "powershell.prop_add" => super::operators::emit_prop_add(chunks, current, line),
        "powershell.collection_clear" => {
            super::operators::emit_collection_clear(chunks, current, line)
        }
        "powershell.count" => super::operators::emit_count(chunks, current, line),
        "powershell.collection_remove" => {
            super::operators::emit_collection_remove(chunks, current, line)
        }
        "powershell.compare_object" => {
            super::operators::emit_compare_object(chunks, current, line)
        }
        "powershell.properties" => {
            super::operators::emit_psobject_properties(chunks, current, argc, line)
        }
        "powershell.splat_arg" => super::operators::emit_splat_arg(chunks, current, line),
        "powershell.unwrap_single" => super::operators::emit_unwrap_single(chunks, current, line),
        "powershell.out_null" => super::operators::emit_out_null(chunks, current, line),
        "powershell.index_get" => super::operators::emit_index_get(chunks, current, line),
        "powershell.index_set" => super::operators::emit_index_set(chunks, current, line),
        "powershell.member_dyn" => super::operators::emit_member_dyn(chunks, current, line),
        "powershell.to_int" => super::operators::emit_to_int(chunks, current, line),
        "powershell.to_char" => super::operators::emit_to_char(chunks, current, line),
        "powershell.compare_to" => super::operators::emit_compare_to(chunks, current, line),
        "powershell.eq" => super::operators::emit_case_folding_eq(chunks, current, false, line),
        "powershell.ne" => super::operators::emit_case_folding_eq(chunks, current, true, line),
        "powershell.get_type" => super::operators::emit_get_type(chunks, current, line),
        "powershell.to_display" => super::display::emit_to_display(chunks, current, line),
        "powershell.format_list" => {
            super::display::emit_format(chunks, current, super::display::FormatMode::List, line)
        }
        "powershell.format_table" => {
            super::display::emit_format(chunks, current, super::display::FormatMode::Table, line)
        }
        "powershell.format_table_no_headers" => super::display::emit_format(
            chunks,
            current,
            super::display::FormatMode::TableNoHeaders,
            line,
        ),
        "powershell.format_wide" => {
            super::display::emit_format(chunks, current, super::display::FormatMode::Wide, line)
        }
        "powershell.to_json" => super::json::emit_to_json(chunks, current, argc, line),
        "powershell.from_json" => super::json::emit_from_json(chunks, current, argc, line),
        _ => return false,
    }
    true
}
