use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;
use std::sync::Arc;
use vybe_runtime::value::Value;

use vybe_ast::{ExprKind, Expression, Literal, ObjectProperty};

use super::build::{expr, str_lit};

/// `SDL_CreateRGBSurface(flags, w, h, depth, rmask, gmask, bmask, amask)`
/// → an offscreen surface the GUEST owns.
///
/// Built at the AST level rather than as raw bytecode: constructing an object
/// in chunk emission is the wrong tool and far more code. `posix_adapter`
/// already returns `ExprKind::Object` results the same way.
///
/// Shape mirrors the fields a software renderer touches:
///
/// ```text
/// { w, h, depth, pixels: [], pitch, format: { palette: [] } }
/// ```
///
/// `pixels` starts EMPTY and grows as the renderer writes — Doom rewrites every
/// pixel each frame, and a runtime-sized zero-fill would need a length the
/// declaration cannot see. Reads of never-written pixels degrade to palette
/// entry 0 at the host rather than faulting.
pub fn create_rgb_surface(
    w: Expression,
    h: Expression,
    depth: Expression,
    pitch: Expression,
) -> Expression {
    let kv = |k: &str, v: Expression| ObjectProperty::KeyValue { key: str_lit(k), value: v };
    let empty = || expr(ExprKind::Array(Vec::new()));
    expr(ExprKind::Object(vec![
        kv("w", w),
        kv("h", h),
        kv("depth", depth),
        kv("pitch", pitch),
        kv("pixels", empty()),
        kv(
            "format",
            expr(ExprKind::Object(vec![
                kv("palette", empty()),
                kv(
                    "BytesPerPixel",
                    expr(ExprKind::Lit(Literal::Int(1))),
                ),
            ])),
        ),
    ]))
}

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

fn emit_string_concat(chunks: &mut [Chunk], current: usize, left: u16, right: u16, line: u32) {
    let concat_idx = chunks[current].add_import("ecma:string", "concat");
    emit_get_local(chunks, current, left, line);
    emit_get_local(chunks, current, right, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunks[current].emit(2u8, line);
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

fn emit_load_f64_from_struct(chunks: &mut [Chunk], current: usize, ptr_slot: u16, field: &str, line: u32) {
    emit_get_local(chunks, current, ptr_slot, line);
    let field_key = chunks[current].add_constant(Value::String(Arc::from(field)));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, field_key, line);
}

fn emit_cstring_to_text(
    chunks: &mut [Chunk],
    current: usize,
    ptr_slot: u16,
    out_slot: u16,
    _idx_slot: u16,
    _byte_slot: u16,
    line: u32,
) {
    // In Vybe C, string literals are passed directly as Vybe String values.
    // For now, assume it's a string literal and just copy it over.
    emit_get_local(chunks, current, ptr_slot, line);
    emit_set_local(chunks, current, out_slot, line);
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

/// A channel value POSITIONED into a packed 0xAARRGGBB colour: `(v & 0xFF) << shift`.
///
/// Distinct from `emit_u8_from_u32_slot`, which is the inverse — it EXTRACTS a
/// channel with `(v >> shift) & 0xFF`. `emit_pack_color` used the extractor to
/// pack, so for any channel below 256 the shift produced 0 and only blue (shift
/// 0) survived: every colour arrived at the host as pure blue.
fn emit_u8_into_u32_slot(chunks: &mut [Chunk], current: usize, slot: u16, shift: u8, line: u32) {
    emit_get_local(chunks, current, slot, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    chunks[current].emit_i32_const(255, line);
    chunks[current].emit_op(Op::I32_AND, line);
    if shift > 0 {
        chunks[current].emit_i32_const(shift.into(), line);
        chunks[current].emit_op(Op::I32_SHL, line);
    }
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
    emit_u8_into_u32_slot(chunks, current, r_slot, 16, line);
    emit_u8_into_u32_slot(chunks, current, g_slot, 8, line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_u8_into_u32_slot(chunks, current, b_slot, 0, line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_u8_into_u32_slot(chunks, current, a_slot, 24, line);
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
    let surface_name = chunks[current].alloc_scratch(1);
    let suffix = chunks[current].alloc_scratch(1);

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

    // SDL window owns a dedicated Canvas control so SDL surface calls render to a
    // real canvas widget (not the form-overlay path).
    emit_get_local(chunks, current, window, line);
    chunks[current].emit_string_const("_surface", line);
    emit_set_local(chunks, current, suffix, line);
    emit_string_concat(chunks, current, window, suffix, line);
    emit_set_local(chunks, current, surface_name, line);

    emit_get_local(chunks, current, window, line);
    chunks[current].emit_string_const("Canvas", line);
    emit_get_local(chunks, current, surface_name, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_i32_const(0, line);
    emit_get_local(chunks, current, w, line);
    emit_get_local(chunks, current, h, line);
    emit_gui_call(chunks, current, "addControl", 7, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_get_local(chunks, current, window, line);
    chunks[current].emit_string_const("sdl_surface", line);
    emit_get_local(chunks, current, surface_name, line);
    emit_gui_call(chunks, current, "setProperty", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_set_control_property(chunks, current, surface_name, "width", w, line);
    emit_set_control_property(chunks, current, surface_name, "height", h, line);
    emit_set_control_property(chunks, current, surface_name, "left", x, line);
    emit_set_control_property(chunks, current, surface_name, "top", y, line);

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
    // DERIVE the surface control's name (`<window>_surface`) the same way
    // `emit_sdl_create_window` builds it, rather than reading back an
    // `sdl_surface` property.
    //
    // The property round-trip returned EMPTY: `setProperty`/`getProperty` are
    // control-oriented and the window is a FORM, so nothing was stored. Every
    // draw then arrived at the host with an empty control name and was recorded
    // into a canvas belonging to no widget — the window came up blank while
    // every drawing call reported success.
    let window = chunks[current].alloc_scratch(1);
    let suffix = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, window, line);
    chunks[current].emit_string_const("_surface", line);
    emit_set_local(chunks, current, suffix, line);
    emit_string_concat(chunks, current, window, suffix, line);
}

/// `SDL_BlitPaletted(surface, pixels, w, h, palette [, dstW, dstH])`
///
/// The whole graphics requirement of a software renderer. The GUEST owns the
/// pixel buffer — Doom writes straight into `screenbuffer->pixels` — so this
/// forwards it as-is and the host does the palette expansion natively.
/// Trailing destination size is optional; the host defaults it to the source
/// size, so the 5-argument form is a 1:1 blit.
pub fn emit_sdl_blit_paletted(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots: Vec<u16> = (0..argc).map(|_| chunks[current].alloc_scratch(1)).collect();
    // Arguments arrive on the stack in order, so pop them back to front.
    for &slot in slots.iter().rev() {
        emit_set_local(chunks, current, slot, line);
    }
    for &slot in slots.iter() {
        emit_get_local(chunks, current, slot, line);
    }
    emit_gui_call(chunks, current, "sdlBlitPaletted", argc, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_fill_rect(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let surface = chunks[current].alloc_scratch(1);
    let rect = chunks[current].alloc_scratch(1);
    let color = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, color, line);
    emit_set_local(chunks, current, rect, line);
    emit_set_local(chunks, current, surface, line);

    // Call sdlFillRect(surface, rect, color)
    emit_get_local(chunks, current, surface, line);
    emit_get_local(chunks, current, rect, line);
    emit_get_local(chunks, current, color, line);
    emit_gui_call(chunks, current, "sdlFillRect", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_draw_line(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let surface = chunks[current].alloc_scratch(1);
    let x1 = chunks[current].alloc_scratch(1);
    let y1 = chunks[current].alloc_scratch(1);
    let x2 = chunks[current].alloc_scratch(1);
    let y2 = chunks[current].alloc_scratch(1);
    let color = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, color, line);
    emit_set_local(chunks, current, y2, line);
    emit_set_local(chunks, current, x2, line);
    emit_set_local(chunks, current, y1, line);
    emit_set_local(chunks, current, x1, line);
    emit_set_local(chunks, current, surface, line);

    emit_get_local(chunks, current, surface, line);
    emit_get_local(chunks, current, x1, line);
    emit_get_local(chunks, current, y1, line);
    emit_get_local(chunks, current, x2, line);
    emit_get_local(chunks, current, y2, line);
    emit_get_local(chunks, current, color, line);
    emit_gui_call(chunks, current, "sdlDrawLine", 6, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_draw_text(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let surface = chunks[current].alloc_scratch(1);
    let text = chunks[current].alloc_scratch(1);
    let x = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);
    let context = chunks[current].alloc_scratch(1);
    let text_value = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let ch = chunks[current].alloc_scratch(1);

    let color = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, color, line);
    emit_set_local(chunks, current, y, line);
    emit_set_local(chunks, current, x, line);
    emit_set_local(chunks, current, text, line);
    emit_set_local(chunks, current, surface, line);

    emit_get_local(chunks, current, surface, line);
    emit_gui_call(chunks, current, "getContext", 1, line);
    emit_set_local(chunks, current, context, line);

    // Text had NO colour of its own: it inherited whatever fill colour the
    // last FillRect set, so a dark theme drew black text on a black panel and
    // the labels vanished. `SDL_DrawText` is an adapter convenience (real SDL
    // uses SDL_ttf), so it takes the colour explicitly.
    emit_get_local(chunks, current, context, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 16, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 8, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 0, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 24, line);
    emit_gui_call(chunks, current, "canvasSetFillColor", 5, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_get_local(chunks, current, context, line);
    emit_cstring_to_text(
        chunks,
        current,
        text,
        text_value,
        idx,
        ch,
        line,
    );
    emit_get_local(chunks, current, text_value, line);
    emit_get_local(chunks, current, x, line);
    emit_get_local(chunks, current, y, line);
    emit_gui_call(chunks, current, "canvasFillText", 4, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_update_window_surface(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let window = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, window, line);

    // Frame boundary: mark the window's surface so the NEXT draw starts a
    // fresh recording. SDL never clears its surface — the program simply
    // redraws — so without this an animated program appends a whole frame of
    // commands per frame forever and every frame paints over the last.
    let suffix = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("_surface", line);
    emit_set_local(chunks, current, suffix, line);
    emit_string_concat(chunks, current, window, suffix, line);
    emit_gui_call(chunks, current, "sdlPresent", 1, line);
    chunks[current].emit_op(Op::DROP, line);

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

pub fn emit_sdl(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    match name {
        "sdl.SDL_Init" | "libc.sdl.SDL_Init" => {
            emit_sdl_init(chunks, current, argc, line);
            true
        }
        "sdl.SDL_InitSubSystem" | "libc.sdl.SDL_InitSubSystem" => {
            emit_sdl_init_subsystem(chunks, current, argc, line);
            true
        }
        "sdl.SDL_Quit" | "libc.sdl.SDL_Quit" => {
            emit_sdl_quit(chunks, current, line);
            true
        }
        "sdl.SDL_CreateWindow" | "libc.sdl.SDL_CreateWindow" => {
            emit_sdl_create_window(chunks, current, argc, line);
            true
        }
        "sdl.SDL_DestroyWindow" | "libc.sdl.SDL_DestroyWindow" => {
            emit_sdl_destroy_window(chunks, current, argc, line);
            true
        }
        "sdl.SDL_GetWindowSurface" | "libc.sdl.SDL_GetWindowSurface" => {
            emit_sdl_get_window_surface(chunks, current, argc, line);
            true
        }
        "sdl.SDL_BlitPaletted" | "libc.sdl.SDL_BlitPaletted" => {
            emit_sdl_blit_paletted(chunks, current, argc, line);
            true
        }
        "sdl.SDL_FillRect" | "libc.sdl.SDL_FillRect" => {
            emit_sdl_fill_rect(chunks, current, argc, line);
            true
        }
        "sdl.SDL_DrawLine" | "libc.sdl.SDL_DrawLine" => {
            emit_sdl_draw_line(chunks, current, argc, line);
            true
        }
        "sdl.SDL_UpdateWindowSurface" | "libc.sdl.SDL_UpdateWindowSurface" => {
            emit_sdl_update_window_surface(chunks, current, argc, line);
            true
        }
        "sdl.SDL_DrawText" | "libc.sdl.SDL_DrawText" => {
            emit_sdl_draw_text(chunks, current, argc, line);
            true
        }
        "sdl.SDL_Delay" | "libc.sdl.SDL_Delay" => {
            emit_sdl_delay(chunks, current, argc, line);
            true
        }
        "sdl.SDL_MapRGB" | "libc.sdl.SDL_MapRGB" => {
            emit_sdl_map_rgb(chunks, current, argc, line);
            true
        }
        "sdl.SDL_MapRGBA" | "libc.sdl.SDL_MapRGBA" => {
            emit_sdl_map_rgba(chunks, current, argc, line);
            true
        }
        "sdl.SDL_ShowWindow" | "libc.sdl.SDL_ShowWindow" => {
            emit_sdl_show_window(chunks, current, argc, line);
            true
        }
        "sdl.SDL_HideWindow" | "libc.sdl.SDL_HideWindow" => {
            emit_sdl_hide_window(chunks, current, argc, line);
            true
        }
        "sdl.SDL_ShowSimpleMessageBox" | "libc.sdl.SDL_ShowSimpleMessageBox" => {
            emit_sdl_show_simple_message_box(chunks, current, argc, line);
            true
        }
        _ => false,
    }
}
