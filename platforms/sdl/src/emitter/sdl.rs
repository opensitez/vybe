use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

fn emit_gui_call(chunks: &mut [Chunk], current: usize, func: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("vybe:gui", func);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_zero_i32(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
}

fn emit_set_local(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_get_local(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn emit_drop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
}

fn emit_number_from_property(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    property: &str,
    line: u32,
) {
    emit_get_local(chunks, current, slot, line);
    chunks[current].emit_string_const(property, line);
    emit_gui_call(chunks, current, "getProperty", 2, line);

    let num_idx = chunks[current].add_import("ecma:number", "Number");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, num_idx, line);
    chunks[current].emit(1, line);
}

fn emit_set_control_property(
    chunks: &mut [Chunk],
    current: usize,
    control_slot: u16,
    property: &str,
    value_slot: u16,
    line: u32,
) {
    emit_get_local(chunks, current, control_slot, line);
    chunks[current].emit_string_const(property, line);
    emit_get_local(chunks, current, value_slot, line);
    emit_gui_call(chunks, current, "setProperty", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_set_control_property_bool_string(
    chunks: &mut [Chunk],
    current: usize,
    control_slot: u16,
    property: &str,
    value: &str,
    line: u32,
) {
    emit_get_local(chunks, current, control_slot, line);
    chunks[current].emit_string_const(property, line);
    chunks[current].emit_string_const(value, line);
    emit_gui_call(chunks, current, "setProperty", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_u8_from_u32_slot(chunks: &mut [Chunk], current: usize, slot: u16, shift: u8, line: u32) {
    emit_get_local(chunks, current, slot, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    if shift > 0 {
        chunks[current].emit_i32_const(shift.into(), line);
        chunks[current].emit_op(Op::I32_SHR_U, line);
    }
    chunks[current].emit_i32_const(255, line);
    chunks[current].emit_op(Op::I32_AND, line);
}

fn emit_u8_from_u32_slot_f64(chunks: &mut [Chunk], current: usize, slot: u16, shift: u8, line: u32) {
    emit_u8_from_u32_slot(chunks, current, slot, shift, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
}

fn emit_pack_color(
    chunks: &mut [Chunk],
    current: usize,
    r_slot: u16,
    g_slot: u16,
    b_slot: u16,
    a_slot: u16,
    line: u32,
) {
    emit_u8_from_u32_slot(chunks, current, r_slot, 16, line);
    emit_u8_from_u32_slot(chunks, current, g_slot, 8, line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_u8_from_u32_slot(chunks, current, b_slot, 0, line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_u8_from_u32_slot(chunks, current, a_slot, 24, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

pub fn emit_sdl_init(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_drop(chunks, current, argc, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_init_subsystem(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_drop(chunks, current, argc, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_quit(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_create_window(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let window = chunks[current].alloc_scratch(1);
    let title = chunks[current].alloc_scratch(1);
    let x = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);
    let w = chunks[current].alloc_scratch(1);
    let h = chunks[current].alloc_scratch(1);
    let _flags = chunks[current].alloc_scratch(1);

    // SDL_CreateWindow(title, x, y, w, h, flags)
    emit_set_local(chunks, current, _flags, line);
    emit_set_local(chunks, current, h, line);
    emit_set_local(chunks, current, w, line);
    emit_set_local(chunks, current, y, line);
    emit_set_local(chunks, current, x, line);
    emit_set_local(chunks, current, title, line);

    emit_get_local(chunks, current, title, line);
    emit_gui_call(chunks, current, "createForm", 1, line);
    emit_set_local(chunks, current, window, line);

    emit_set_control_property(chunks, current, window, "left", x, line);
    emit_set_control_property(chunks, current, window, "top", y, line);
    emit_set_control_property(chunks, current, window, "width", w, line);
    emit_set_control_property(chunks, current, window, "height", h, line);

    emit_get_local(chunks, current, window, line);
}

pub fn emit_sdl_destroy_window(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let window = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, window, line);
    emit_get_local(chunks, current, window, line);
    emit_gui_call(chunks, current, "closeForm", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_get_window_surface(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let window = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, window, line);
    emit_get_local(chunks, current, window, line);
}

pub fn emit_sdl_fill_rect(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let surface = chunks[current].alloc_scratch(1);
    let _rect = chunks[current].alloc_scratch(1);
    let color = chunks[current].alloc_scratch(1);
    let context = chunks[current].alloc_scratch(1);
    let width = chunks[current].alloc_scratch(1);
    let height = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, color, line);
    emit_set_local(chunks, current, _rect, line);
    emit_set_local(chunks, current, surface, line);

    emit_number_from_property(chunks, current, surface, "width", line);
    emit_set_local(chunks, current, width, line);
    emit_number_from_property(chunks, current, surface, "height", line);
    emit_set_local(chunks, current, height, line);

    emit_get_local(chunks, current, surface, line);
    emit_gui_call(chunks, current, "getContext", 1, line);
    emit_set_local(chunks, current, context, line);

    emit_get_local(chunks, current, context, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 16, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 8, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 0, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 24, line);
    emit_gui_call(chunks, current, "canvasSetFillColor", 5, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_get_local(chunks, current, context, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_f64_const(0.0, line);
    emit_get_local(chunks, current, width, line);
    emit_get_local(chunks, current, height, line);
    emit_gui_call(chunks, current, "canvasFillRect", 5, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_update_window_surface(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let window = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, window, line);
    emit_get_local(chunks, current, window, line);
    emit_gui_call(chunks, current, "runApplication", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_delay(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let delay_ms = chunks[current].alloc_scratch(1);
    let sub_idx = chunks[current].add_import("wasi:clocks/monotonic-clock", "subscribe-duration");
    let block_idx = chunks[current].add_import("wasi:io/poll", "[method]pollable.block");

    emit_set_local(chunks, current, delay_ms, line);
    emit_get_local(chunks, current, delay_ms, line);
    chunks[current].emit_f64_const(1_000_000.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, sub_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, block_idx, line);
    chunks[current].emit(1, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_show_window(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let window = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, window, line);
    emit_set_control_property_bool_string(chunks, current, window, "visible", "true", line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_hide_window(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let window = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, window, line);
    emit_set_control_property_bool_string(chunks, current, window, "visible", "false", line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_show_simple_message_box(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let _flags = chunks[current].alloc_scratch(1);
    let title = chunks[current].alloc_scratch(1);
    let text = chunks[current].alloc_scratch(1);
    let _window = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, _window, line);
    emit_set_local(chunks, current, text, line);
    emit_set_local(chunks, current, title, line);
    emit_set_local(chunks, current, _flags, line);

    emit_get_local(chunks, current, title, line);
    emit_get_local(chunks, current, text, line);
    emit_gui_call(chunks, current, "msgBox", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_map_rgb(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let fmt = chunks[current].alloc_scratch(1);
    let r = chunks[current].alloc_scratch(1);
    let g = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, b, line);
    emit_set_local(chunks, current, g, line);
    emit_set_local(chunks, current, r, line);
    emit_set_local(chunks, current, fmt, line);

    chunks[current].emit_i32_const(255, line);
    emit_set_local(chunks, current, a, line);
    emit_pack_color(chunks, current, r, g, b, a, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
}

pub fn emit_sdl_map_rgba(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let fmt = chunks[current].alloc_scratch(1);
    let r = chunks[current].alloc_scratch(1);
    let g = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, a, line);
    emit_set_local(chunks, current, b, line);
    emit_set_local(chunks, current, g, line);
    emit_set_local(chunks, current, r, line);
    emit_set_local(chunks, current, fmt, line);

    emit_pack_color(chunks, current, r, g, b, a, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
}
