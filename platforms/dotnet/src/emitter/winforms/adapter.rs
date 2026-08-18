//! Shared WinForms-to-vybe GUI adapter leaves.

use vybe_runtime::Chunk;

fn emit_gui_call(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("vybe:gui", name);
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, argc, line);
}

pub fn emit_message_box_show(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(chunks, current, "msgBox", argc, line);
}

pub fn emit_application_run(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(
        chunks,
        current,
        vybe_compiler::primitives::gui::HOST_FN_RUN_APPLICATION,
        argc,
        line,
    );
}

pub fn emit_application_exit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(
        chunks,
        current,
        vybe_compiler::primitives::gui::HOST_FN_APP_EXIT,
        argc,
        line,
    );
}

pub fn emit_noop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(chunks, current, "noop", argc, line);
}

