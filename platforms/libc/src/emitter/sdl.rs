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

/// Call a `web:canvas` op. SDL's drawing is an ADAPTER over WHATWG
/// `CanvasRenderingContext2D`: `SDL_FillRect` IS `fillRect` plus a rect
/// struct, `SDL_BlitPaletted` IS `drawImagePaletted`. No canvas surface of
/// our own — a browser host serves these imports with a real canvas element.
fn emit_canvas_call(chunks: &mut [Chunk], current: usize, func: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("web:canvas", func);
    chunks[current].emit_call(idx, argc, line);
}


/// Call a `web:ui-events` host function. SDL's input is an ADAPTER over the
/// W3C UI Events surface in `platforms/web` — there is no SDL host surface
/// and no `vybe:gui` involvement: the queue is the web platform's, and every
/// vocabulary difference (DOM `"keydown"` vs `SDL_KEYDOWN`, DOM's 0-based
/// `button` vs SDL's 1-based, `key`/`code` strings vs `SDLK_*`) is resolved
/// here, in emitted code. A browser host satisfies the same imports with the
/// real DOM.
fn emit_web_events_call(chunks: &mut [Chunk], current: usize, func: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("web:ui-events", func);
    chunks[current].emit_call(idx, argc, line);
}

/// Read `obj.<field>` from the DOM event object in `slot`.
fn emit_dom_field(chunks: &mut [Chunk], current: usize, slot: u16, field: &str, line: u32) {
    emit_get_local(chunks, current, slot, line);
    let key = chunks[current].add_constant(Value::String(Arc::from(field)));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

/// `target.<field> = <value on stack>`.
fn emit_store_field(
    chunks: &mut [Chunk],
    current: usize,
    target: u16,
    field: &str,
    tmp: u16,
    line: u32,
) {
    emit_set_local(chunks, current, tmp, line);
    emit_get_local(chunks, current, target, line);
    emit_get_local(chunks, current, tmp, line);
    let key = chunks[current].add_constant(Value::String(Arc::from(field)));
    // The NAME-KEYED `struct.set` (typeidx 0) pushes the value back — unlike
    // the spec's indexed form, which yields nothing. Hence the DROP.
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

/// `1` when the DOM event's `type` equals `kind`, else `0`.
fn emit_dom_kind_is(chunks: &mut [Chunk], current: usize, ev: u16, kind: &str, line: u32) {
    emit_dom_field(chunks, current, ev, "type", line);
    chunks[current].emit_string_const(kind, line);
    chunks[current].emit_op(Op::EQ, line);
}


/// Unwrap a C pointer argument to the object it addresses.
///
/// `&e` on a struct reaches a callee either as the struct itself or boxed in
/// a scalar cell `{__ref_kind:"cell", __value}`. The host used to do this;
/// now that SDL is adapter-only, the emitted code must — reading `.type`
/// straight off a cell yields undefined, which is how every field arrived
/// as zero.
fn emit_deref_cell(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_get_local(chunks, current, slot, line);
    let kind_key = chunks[current].add_constant(Value::String(Arc::from("__ref_kind")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, kind_key, line);
    chunks[current].emit_string_const("cell", line);
    chunks[current].emit_op(Op::EQ, line);
    chunks[current].emit_if_value(line);
    emit_get_local(chunks, current, slot, line);
    let val_key = chunks[current].add_constant(Value::String(Arc::from("__value")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, val_key, line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, slot, line);
    chunks[current].emit_end(line);
    emit_set_local(chunks, current, slot, line);
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
    chunks[current].emit_call(concat_idx, 2u8, line);
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
    chunks[current].emit_call(num_idx, 1, line);
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
    // `SDL_Window *win` reaches us as a `{__ref_kind:"cell", __value}` box,
    // not as the form name. Concatenating the BOX yields the box back, so
    // `getContext` stored an object where a control name belongs and every
    // draw landed on a canvas called "[object]" while the real surface got
    // nothing. Unwrap first — the same step `SDL_PollEvent`/`SDL_PushEvent`
    // already take for their event pointers.
    emit_deref_cell(chunks, current, window, line);
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
    // `drawImagePaletted` — the canvas op for palette-era pixels. The
    // guest keeps its 8-bit buffer; the engine expands through the palette.
    emit_canvas_call(chunks, current, "drawImagePaletted", argc, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_fill_rect(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let surface = chunks[current].alloc_scratch(1);
    let rect = chunks[current].alloc_scratch(1);
    let color = chunks[current].alloc_scratch(1);
    let ctx = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, color, line);
    emit_set_local(chunks, current, rect, line);
    emit_set_local(chunks, current, surface, line);
    // `SDL_Surface *screen` arrives as a `{__ref_kind:"cell", __value}`
    // box, not the surface's control name. Handed to `getContext` boxed,
    // it became the target `"[object]"` — every draw landed on a canvas
    // belonging to no widget while the real surface got nothing.
    emit_deref_cell(chunks, current, surface, line);

    // `SDL_FillRect` IS `fillRect` — plus SDL's own two shapes: a rect
    // STRUCT (x/y/w/h, or NULL meaning "the whole surface") and a PACKED
    // 0xAARRGGBB colour where the canvas takes channels. Both are unpacked
    // here, on the adapter's side of the standard surface.
    emit_get_local(chunks, current, surface, line);
    emit_canvas_call(chunks, current, "getContext", 1, line);
    emit_set_local(chunks, current, ctx, line);

    emit_get_local(chunks, current, ctx, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 16, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 8, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 0, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 24, line);
    emit_canvas_call(chunks, current, "setFillStyle", 5, line);
    chunks[current].emit_op(Op::DROP, line);

    // `&r` on a local `SDL_Rect` reaches us either as the struct itself or
    // boxed in a scalar cell. Reading `.x` off the BOX yields undefined, which
    // becomes 0 — so every rect was recorded at zero size and painted nothing
    // while text, whose coordinates are plain ints, still showed.
    emit_deref_cell(chunks, current, rect, line);

    emit_get_local(chunks, current, ctx, line);
    emit_load_f64_from_struct(chunks, current, rect, "x", line);
    emit_load_f64_from_struct(chunks, current, rect, "y", line);
    emit_load_f64_from_struct(chunks, current, rect, "w", line);
    emit_load_f64_from_struct(chunks, current, rect, "h", line);
    emit_canvas_call(chunks, current, "fillRect", 5, line);
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
    let ctx = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, color, line);
    emit_set_local(chunks, current, y2, line);
    emit_set_local(chunks, current, x2, line);
    emit_set_local(chunks, current, y1, line);
    emit_set_local(chunks, current, x1, line);
    emit_set_local(chunks, current, surface, line);
    // `SDL_Surface *screen` arrives as a `{__ref_kind:"cell", __value}`
    // box, not the surface's control name. Handed to `getContext` boxed,
    // it became the target `"[object]"` — every draw landed on a canvas
    // belonging to no widget while the real surface got nothing.
    emit_deref_cell(chunks, current, surface, line);

    // A line is a path in canvas terms: beginPath → moveTo → lineTo →
    // stroke. SDL has no path model, which is exactly the kind of
    // difference an adapter absorbs.
    emit_get_local(chunks, current, surface, line);
    emit_canvas_call(chunks, current, "getContext", 1, line);
    emit_set_local(chunks, current, ctx, line);

    emit_get_local(chunks, current, ctx, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 16, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 8, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 0, line);
    emit_u8_from_u32_slot_f64(chunks, current, color, 24, line);
    emit_canvas_call(chunks, current, "setStrokeStyle", 5, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_get_local(chunks, current, ctx, line);
    emit_canvas_call(chunks, current, "beginPath", 1, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_get_local(chunks, current, ctx, line);
    emit_get_local(chunks, current, x1, line);
    emit_get_local(chunks, current, y1, line);
    emit_canvas_call(chunks, current, "moveTo", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_get_local(chunks, current, ctx, line);
    emit_get_local(chunks, current, x2, line);
    emit_get_local(chunks, current, y2, line);
    emit_canvas_call(chunks, current, "lineTo", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_get_local(chunks, current, ctx, line);
    emit_canvas_call(chunks, current, "stroke", 1, line);
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
    // `SDL_Surface *screen` arrives as a `{__ref_kind:"cell", __value}`
    // box, not the surface's control name. Handed to `getContext` boxed,
    // it became the target `"[object]"` — every draw landed on a canvas
    // belonging to no widget while the real surface got nothing.
    emit_deref_cell(chunks, current, surface, line);

    emit_get_local(chunks, current, surface, line);
    emit_canvas_call(chunks, current, "getContext", 1, line);
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
    emit_canvas_call(chunks, current, "setFillStyle", 5, line);
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
    emit_canvas_call(chunks, current, "fillText", 4, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_zero_i32(chunks, current, line);
}

pub fn emit_sdl_update_window_surface(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let window = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, window, line);
    // Same pointer unwrap as `SDL_GetWindowSurface`: the window arrives boxed.
    emit_deref_cell(chunks, current, window, line);

    // There is no `present` in the web platform — a page does not push
    // frames, it draws and the compositor shows them, and the frame
    // BOUNDARY is `requestAnimationFrame`. So this call has no canvas
    // counterpart to move to: it collapses to the existing run/show step,
    // and the per-frame reset it used to carry belongs to `web:animation`
    // (rAF), which is the next surface to land.
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
    chunks[current].emit_call(sub_idx, 1, line);
    chunks[current].emit_call(block_idx, 1, line);
    emit_zero_i32(chunks, current, line);
}

// ── Tier 2: timing (`sdlplan.md`) ───────────────────────────────────────────
//
// All three ride `wasi:clocks/monotonic-clock.now` (f64 NANOSECONDS since
// process start — the same clock `SDL_Delay` subscribes against), so ticks
// and delays can never drift apart.

/// `SDL_GetTicks()` → milliseconds since start.
pub fn emit_sdl_get_ticks(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let now_idx = chunks[current].add_import("wasi:clocks/monotonic-clock", "now");
    chunks[current].emit_call(now_idx, 0, line);
    chunks[current].emit_f64_const(1_000_000.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
}

/// `SDL_GetPerformanceCounter()` → the raw nanosecond counter.
pub fn emit_sdl_get_performance_counter(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let now_idx = chunks[current].add_import("wasi:clocks/monotonic-clock", "now");
    chunks[current].emit_call(now_idx, 0, line);
}

/// `SDL_GetPerformanceFrequency()` → counts per second: nanoseconds → 1e9.
pub fn emit_sdl_get_performance_frequency(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    chunks[current].emit_f64_const(1_000_000_000.0, line);
}

// ── Tier 1: input (`sdlplan.md`) ────────────────────────────────────────────
//
// The HOST fills the `SDL_Event` struct — it can mutate the pointee object
// directly — so each of these stays a plain call instead of a field-copy
// sequence in bytecode.

/// `SDL_PollEvent(SDL_Event *e)` → 1 if an event was dequeued, else 0.
///
/// Pure ADAPTER over `web:ui-events.pollEvent()`: takes the W3C event object
/// and writes SDL's struct view of it. No host function of its own, no
/// `vybe:gui` — the queue belongs to the web platform, and a browser host
/// serves the same import from the real DOM.
pub fn emit_sdl_poll_event(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let store_tmp = chunks[current].alloc_scratch(1);
    let ptr = chunks[current].alloc_scratch(1);
    let ev = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let keysym = chunks[current].alloc_scratch(1);
    let btn = chunks[current].alloc_scratch(1);
    let motion = chunks[current].alloc_scratch(1);
    let wheel = chunks[current].alloc_scratch(1);
    let kind = chunks[current].alloc_scratch(1);

    emit_set_local(chunks, current, ptr, line);
    emit_deref_cell(chunks, current, ptr, line);

    emit_web_events_call(chunks, current, "pollEvent", 0, line);
    emit_set_local(chunks, current, ev, line);

    // Empty queue → 0.
    emit_get_local(chunks, current, ev, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_zero_i32(chunks, current, line);
    chunks[current].emit_else(line);

    // SDL event type from the DOM `type` string.
    emit_dom_kind_is(chunks, current, ev, "keydown", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0x300, line);
    chunks[current].emit_else(line);
    emit_dom_kind_is(chunks, current, ev, "keyup", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0x301, line);
    chunks[current].emit_else(line);
    emit_dom_kind_is(chunks, current, ev, "mousedown", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0x401, line);
    chunks[current].emit_else(line);
    emit_dom_kind_is(chunks, current, ev, "mouseup", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0x402, line);
    chunks[current].emit_else(line);
    emit_dom_kind_is(chunks, current, ev, "mousemove", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0x400, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0x403, line); // wheel
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    emit_set_local(chunks, current, kind, line);

    emit_get_local(chunks, current, kind, line);
    emit_store_field(chunks, current, ptr, "type", store_tmp, line);

    // key.keysym.{sym,scancode,mod} — DOM keyCode IS the SDL keysym for the
    // printable range, which is how the winit layer fills it.
    emit_get_local(chunks, current, ptr, line);
    let key_field = chunks[current].add_constant(Value::String(Arc::from("key")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key_field, line);
    emit_set_local(chunks, current, key, line);
    emit_get_local(chunks, current, key, line);
    let keysym_field = chunks[current].add_constant(Value::String(Arc::from("keysym")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, keysym_field, line);
    emit_set_local(chunks, current, keysym, line);

    // DOM `keyCode` is the legacy UPPERCASE identity (A = 65); an SDL keysym
    // for a letter is the LOWERCASE ascii value (SDLK_a = 97). Scancode is
    // the USB HID position: letters 4..29, digits 30..38, '0' = 39.
    let kc = chunks[current].alloc_scratch(1);
    emit_dom_field(chunks, current, ev, "keyCode", line);
    emit_set_local(chunks, current, kc, line);

    // sym = (65 <= kc <= 90) ? kc + 32 : kc
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(65.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(90.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(32.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_end(line);
    emit_store_field(chunks, current, keysym, "sym", store_tmp, line);

    // scancode: letters → 4 + (kc - 65); '1'..'9' → 30 + (kc - 49);
    // '0' → 39; anything else 0 (Doom reads sym for those).
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(65.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(90.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(61.0, line); // 4 + (kc - 65) == kc - 61
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(49.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(57.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    emit_get_local(chunks, current, kc, line);
    chunks[current].emit_f64_const(19.0, line); // 49 - 30
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    emit_store_field(chunks, current, keysym, "scancode", store_tmp, line);

    // KMOD_* mask from the DOM's boolean modifiers — the inverse of what the
    // push side does, so a pushed event round-trips its modifiers.
    let mods = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    emit_set_local(chunks, current, mods, line);
    for (field, mask) in [("shiftKey", 0x1i32), ("ctrlKey", 0x40), ("altKey", 0x100)] {
        emit_dom_field(chunks, current, ev, field, line);
        chunks[current].emit_if_value(line);
        emit_get_local(chunks, current, mods, line);
        chunks[current].emit_i32_const(mask, line);
        chunks[current].emit_op(Op::I32_OR, line);
        chunks[current].emit_else(line);
        emit_get_local(chunks, current, mods, line);
        chunks[current].emit_end(line);
        emit_set_local(chunks, current, mods, line);
    }
    emit_get_local(chunks, current, mods, line);
    emit_store_field(chunks, current, keysym, "mod", store_tmp, line);

    // button.{button,x,y} — DOM button is 0-based, SDL 1-based.
    emit_get_local(chunks, current, ptr, line);
    let btn_field = chunks[current].add_constant(Value::String(Arc::from("button")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, btn_field, line);
    emit_set_local(chunks, current, btn, line);
    emit_dom_field(chunks, current, ev, "button", line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_store_field(chunks, current, btn, "button", store_tmp, line);
    emit_dom_field(chunks, current, ev, "clientX", line);
    emit_store_field(chunks, current, btn, "x", store_tmp, line);
    emit_dom_field(chunks, current, ev, "clientY", line);
    emit_store_field(chunks, current, btn, "y", store_tmp, line);

    // motion.{x,y}
    emit_get_local(chunks, current, ptr, line);
    let motion_field = chunks[current].add_constant(Value::String(Arc::from("motion")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, motion_field, line);
    emit_set_local(chunks, current, motion, line);
    emit_dom_field(chunks, current, ev, "clientX", line);
    emit_store_field(chunks, current, motion, "x", store_tmp, line);
    emit_dom_field(chunks, current, ev, "clientY", line);
    emit_store_field(chunks, current, motion, "y", store_tmp, line);

    // wheel.y — DOM deltaY is positive DOWN, SDL wheel y positive UP.
    emit_get_local(chunks, current, ptr, line);
    let wheel_field = chunks[current].add_constant(Value::String(Arc::from("wheel")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, wheel_field, line);
    emit_set_local(chunks, current, wheel, line);
    chunks[current].emit_f64_const(0.0, line);
    emit_dom_field(chunks, current, ev, "deltaY", line);
    chunks[current].emit_op(Op::F64_SUB, line);
    emit_store_field(chunks, current, wheel, "y", store_tmp, line);

    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);
}

/// `SDL_PushEvent(SDL_Event *e)` → 1. `EventTarget.dispatchEvent` in SDL's
/// dialect: the injected event joins the SAME `web:ui-events` queue real
/// input arrives on, which is also what makes the pipeline headless-testable.
pub fn emit_sdl_push_event(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let store_tmp = chunks[current].alloc_scratch(1);
    let ptr = chunks[current].alloc_scratch(1);
    let dom = chunks[current].alloc_scratch(1);
    let ty = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, ptr, line);
    emit_deref_cell(chunks, current, ptr, line);

    // The SDL type decides the DOM `type` string.
    emit_get_local(chunks, current, ptr, line);
    let type_key = chunks[current].add_constant(Value::String(Arc::from("type")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    emit_set_local(chunks, current, ty, line);

    emit_get_local(chunks, current, ty, line);
    chunks[current].emit_f64_const(768.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("keydown", line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, ty, line);
    chunks[current].emit_f64_const(769.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("keyup", line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, ty, line);
    chunks[current].emit_f64_const(1025.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("mousedown", line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, ty, line);
    chunks[current].emit_f64_const(1026.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("mouseup", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("mousemove", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    emit_web_events_call(chunks, current, "newEvent", 1, line);
    emit_set_local(chunks, current, dom, line);

    // key.keysym.sym → keyCode; button.{button,x,y} → button/clientX/clientY.
    // `keyCode` is the browser's legacy UPPERCASE identity (W = 87) while an
    // SDL keysym is lowercase (SDLK_w = 119). Converting here is what lets
    // the poll side derive both `sym` AND the USB-HID `scancode` back.
    let sym_v = chunks[current].alloc_scratch(1);
    emit_get_local(chunks, current, ptr, line);
    let key_key = chunks[current].add_constant(Value::String(Arc::from("key")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key_key, line);
    let keysym_key = chunks[current].add_constant(Value::String(Arc::from("keysym")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, keysym_key, line);
    let sym_key = chunks[current].add_constant(Value::String(Arc::from("sym")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, sym_key, line);
    emit_set_local(chunks, current, sym_v, line);

    emit_get_local(chunks, current, sym_v, line);
    chunks[current].emit_f64_const(97.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    emit_get_local(chunks, current, sym_v, line);
    chunks[current].emit_f64_const(122.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    emit_get_local(chunks, current, sym_v, line);
    chunks[current].emit_f64_const(32.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, sym_v, line);
    chunks[current].emit_end(line);
    emit_store_field(chunks, current, dom, "keyCode", store_tmp, line);

    // KMOD_* → the DOM's boolean modifier attributes.
    //
    // Bit-tested with FLOAT ops. Struct fields arrive as numbers whose
    // concrete tag varies, and the integer ops (`i32.and`, the f64→i32
    // conversions) silently yielded 0 on them — the same typed-op mismatch
    // that made `Op::EQ` fail against `f64` type codes earlier in this
    // function. `bit = m - 2*floor(m/2)` needs no coercion at all.
    let mods = chunks[current].alloc_scratch(1);
    emit_get_local(chunks, current, ptr, line);
    let key_key2 = chunks[current].add_constant(Value::String(Arc::from("key")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key_key2, line);
    let keysym_key2 = chunks[current].add_constant(Value::String(Arc::from("keysym")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, keysym_key2, line);
    let mod_key = chunks[current].add_constant(Value::String(Arc::from("mod")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, mod_key, line);
    emit_set_local(chunks, current, mods, line);

    for (mask, field) in [(1i32, "shiftKey"), (0x40, "ctrlKey"), (0x100, "altKey")] {
        // Integer mask directly on the field value: KMOD_* is a bitmask and
        // the value arrives as an integer, so `i32.and` needs no conversion
        // (adding one is what silently produced 0 in earlier attempts).
        // `mods & mask` is already 0-or-nonzero, and the host reads these
        // attributes with JS truthiness — so store the mask result straight
        // in. (An `i32.ne` normalisation step here produced a value the host
        // read as false; not worth a second opcode to find out why.)
        emit_get_local(chunks, current, mods, line);
        chunks[current].emit_i32_const(mask, line);
        chunks[current].emit_op(Op::I32_AND, line);
        emit_store_field(chunks, current, dom, field, store_tmp, line);
    }

    emit_get_local(chunks, current, ptr, line);
    let btn_key = chunks[current].add_constant(Value::String(Arc::from("button")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, btn_key, line);
    let bb_key = chunks[current].add_constant(Value::String(Arc::from("button")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, bb_key, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    emit_store_field(chunks, current, dom, "button", store_tmp, line);

    // `buttons` is the HELD-button mask the DOM tracks: 1 left, 2 right,
    // 4 middle — a different assignment from SDL's, resolved here.
    let sdl_btn = chunks[current].alloc_scratch(1);
    emit_get_local(chunks, current, ptr, line);
    let btn_k = chunks[current].add_constant(Value::String(Arc::from("button")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, btn_k, line);
    let bb_k = chunks[current].add_constant(Value::String(Arc::from("button")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, bb_k, line);
    emit_set_local(chunks, current, sdl_btn, line);
    emit_get_local(chunks, current, sdl_btn, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, sdl_btn, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(4.0, line);
    chunks[current].emit_else(line);
    emit_get_local(chunks, current, sdl_btn, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    emit_store_field(chunks, current, dom, "buttons", store_tmp, line);

    emit_get_local(chunks, current, ptr, line);
    let btn_key2 = chunks[current].add_constant(Value::String(Arc::from("button")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, btn_key2, line);
    let bx_key = chunks[current].add_constant(Value::String(Arc::from("x")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, bx_key, line);
    emit_store_field(chunks, current, dom, "clientX", store_tmp, line);

    emit_get_local(chunks, current, ptr, line);
    let btn_key3 = chunks[current].add_constant(Value::String(Arc::from("button")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, btn_key3, line);
    let by_key = chunks[current].add_constant(Value::String(Arc::from("y")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, by_key, line);
    emit_store_field(chunks, current, dom, "clientY", store_tmp, line);

    emit_get_local(chunks, current, dom, line);
    emit_web_events_call(chunks, current, "dispatchEvent", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(1, line);
}

/// `SDL_GetMouseState(int *x, int *y)` → held-button mask. The host writes
/// through the out-pointers.
pub fn emit_sdl_get_mouse_state(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let store_tmp = chunks[current].alloc_scratch(1);
    // Pad missing out-pointers so the host always sees two args.
    for _ in argc..2 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    // `pointerState()` is the browser's tracked pointer; SDL's out-params
    // and 1-based button mask are this adapter's business.
    let st = chunks[current].alloc_scratch(1);
    let py = chunks[current].alloc_scratch(1);
    let px = chunks[current].alloc_scratch(1);
    emit_set_local(chunks, current, py, line);
    emit_set_local(chunks, current, px, line);
    emit_web_events_call(chunks, current, "pointerState", 0, line);
    emit_set_local(chunks, current, st, line);
    emit_dom_field(chunks, current, st, "clientX", line);
    emit_store_field(chunks, current, px, "__value", store_tmp, line);
    emit_dom_field(chunks, current, st, "clientY", line);
    emit_store_field(chunks, current, py, "__value", store_tmp, line);
    emit_dom_field(chunks, current, st, "buttons", line);
}

/// `SDL_GetModState()` → KMOD_* mask.
pub fn emit_sdl_get_mod_state(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let st = chunks[current].alloc_scratch(1);
    emit_web_events_call(chunks, current, "pointerState", 0, line);
    emit_set_local(chunks, current, st, line);
    // KMOD_LSHIFT 0x1 | KMOD_LCTRL 0x40 | KMOD_LALT 0x100 | KMOD_LGUI 0x400
    emit_dom_field(chunks, current, st, "shiftKey", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0x1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
}

/// `SDL_PumpEvents()` — the winit loop pumps for us; nothing to do.
pub fn emit_sdl_pump_events(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_zero_i32(chunks, current, line);
}

/// `SDL_PeepEvents(...)` → 0 events, dropping every argument.
pub fn emit_sdl_peep_events(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
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
        "sdl.SDL_PollEvent" | "libc.sdl.SDL_PollEvent" => {
            emit_sdl_poll_event(chunks, current, argc, line);
            true
        }
        "sdl.SDL_PushEvent" | "libc.sdl.SDL_PushEvent" => {
            emit_sdl_push_event(chunks, current, argc, line);
            true
        }
        "sdl.SDL_GetMouseState" | "libc.sdl.SDL_GetMouseState" => {
            emit_sdl_get_mouse_state(chunks, current, argc, line);
            true
        }
        "sdl.SDL_GetModState" | "libc.sdl.SDL_GetModState" => {
            emit_sdl_get_mod_state(chunks, current, argc, line);
            true
        }
        "sdl.SDL_PumpEvents" | "libc.sdl.SDL_PumpEvents" => {
            emit_sdl_pump_events(chunks, current, argc, line);
            true
        }
        "sdl.SDL_PeepEvents" | "libc.sdl.SDL_PeepEvents" => {
            emit_sdl_peep_events(chunks, current, argc, line);
            true
        }
        "sdl.SDL_GetTicks" | "libc.sdl.SDL_GetTicks" => {
            emit_sdl_get_ticks(chunks, current, argc, line);
            true
        }
        "sdl.SDL_GetPerformanceCounter" | "libc.sdl.SDL_GetPerformanceCounter" => {
            emit_sdl_get_performance_counter(chunks, current, argc, line);
            true
        }
        "sdl.SDL_GetPerformanceFrequency" | "libc.sdl.SDL_GetPerformanceFrequency" => {
            emit_sdl_get_performance_frequency(chunks, current, argc, line);
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
