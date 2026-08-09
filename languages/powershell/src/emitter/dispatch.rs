//! `common:powershell.*` dispatch.
//!
//! Two arms, both proven. Every other PowerShell surface reaches a shared
//! primitive through the profile; these are the only places where none of them
//! can answer, which is the bar `documentation/powershellplan.md` §6.5 sets for
//! adding an emitter at all. Adding a third arm means proving it again.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "powershell.add" => super::operators::emit_add(chunks, current, line),
        "powershell.ensure_array" => {
            super::operators::emit_ensure_array(chunks, current, argc, line)
        }
        // Routes to `primitives::json`, which has no `common:json.*` dispatch
        // key — go and pascal reach it the same way.
        // `[builtin_slots.string] to_string` — the interpolation/concat slot.
        "powershell.to_display" => super::display::emit_to_display(chunks, current, line),
        "powershell.to_json" => super::json::emit_to_json(chunks, current, argc, line),
        "powershell.from_json" => super::json::emit_from_json(chunks, current, argc, line),
        _ => return false,
    }
    true
}
