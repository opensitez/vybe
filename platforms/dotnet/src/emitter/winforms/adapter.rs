//! Shared WinForms-to-vybe GUI adapter leaves.

use vybe_bytecode::{Chunk, opcode::Op};

fn emit_gui_call(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("vybe:gui", name);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

pub fn emit_message_box_show(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(chunks, current, "msgBox", argc, line);
}

pub fn emit_application_run(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(
        chunks,
        current,
        vybe_emitter::gui::HOST_FN_RUN_APPLICATION,
        argc,
        line,
    );
}

pub fn emit_application_exit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(
        chunks,
        current,
        vybe_emitter::gui::HOST_FN_APP_EXIT,
        argc,
        line,
    );
}

pub fn emit_noop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(chunks, current, "noop", argc, line);
}

pub fn emit_control_show(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(chunks, current, "__ctrl_show", argc, line);
}

pub fn emit_control_hide(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(chunks, current, "__ctrl_hide", argc, line);
}

pub fn emit_control_close(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(chunks, current, "__ctrl_close", argc, line);
}

pub fn emit_form_show_dialog(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(chunks, current, "__dlg_showdialog", argc, line);
}

pub fn emit_controls_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gui_call(
        chunks,
        current,
        vybe_emitter::gui::HOST_FN_ADD_CHILD,
        argc,
        line,
    );
}
