//! Offscreen frame capture for GUI programs.
//!
//! A form renders into a single `tiny_skia::Pixmap`, so a screenshot is just
//! that pixmap encoded as PNG — the same rasterise-then-read-bytes path the
//! canvas read-back uses. Two callers:
//!
//! - `--capture <png>` (headless, `gui_launch::capture_gui`) — run the program,
//!   render one frame, write the file, exit. No window is ever created.
//! - `capture` in the step debugger — grab the live frame mid-session.
//!
//! `render_into` is the ONE renderer; `FormApp::render` delegates to it so the
//! windowed and captured frames cannot drift apart.
//!
//! **Everything here reads the DOCUMENT.** A control is
//! `document.createElement(tag)`, its geometry is the laid-out box, and its
//! name is its `id` — so a capture asks the same tree a window paints and the
//! debugger inspects.

use widgets::{FontSystem, Pixmap, RenderContext, SwashCache, fill_background};

/// Paint the document into `pixmap`.
///
/// This is the whole frame: `FormApp::render` calls it for the live window and
/// the capture paths call it for an offscreen buffer.
/// The fonts used to be parameters. They are not any more: the ENGINE owns its
/// own text stack, and a host that supplies one is a host that can supply the
/// wrong one — a frame rendered with fonts other than the ones the engine
/// measured with lays text out to the wrong width.
pub fn render_into(pixmap: &mut Pixmap, scale: f32) {
    fill_background(pixmap, 240, 240, 240, 255);
    // Through `platforms/web`, which forwards to whichever engine is live.
    // This used to call `widgets::dom` directly — around the intermediary
    // rather than through it — so `--engine webcore` swapped the engine and
    // left the renderer pointed at the toolkit's empty tree.
    //
    // `gui_document::active()` rather than `html::active_document()`: the
    // debugger reads from its own REPL thread, where that thread-local is a
    // different, empty document entirely.
    vybe_platform_web::present::render(crate::gui_document::active(), pixmap, scale);
    // ⛔ Nothing composites an overlay on top. A `<canvas>` is an element in
    // the document and paints as part of it, like any other control — so a
    // capture and a window show the same thing BY CONSTRUCTION rather than by
    // agreement. An earlier overlay pass painted drawings that no window ever
    // showed, which made the screenshot say the opposite of the truth.
}

/// Crop `src` to a control's rect. Returns `None` if the rect falls entirely
/// outside the frame — a control that was never laid out.
fn crop(src: &Pixmap, x: f32, y: f32, w: f32, h: f32) -> Option<Pixmap> {
    let x0 = x.max(0.0).round() as u32;
    let y0 = y.max(0.0).round() as u32;
    let x1 = ((x + w).round() as i64).clamp(0, src.width() as i64) as u32;
    let y1 = ((y + h).round() as i64).clamp(0, src.height() as i64) as u32;
    if x1 <= x0 || y1 <= y0 {
        return None;
    }
    let mut out = Pixmap::new(x1 - x0, y1 - y0)?;
    let (sw, ow, oh) = (
        src.width() as usize,
        out.width() as usize,
        out.height() as usize,
    );
    let (sp, op) = (src.data(), out.data_mut());
    for row in 0..oh {
        let s = ((y0 as usize + row) * sw + x0 as usize) * 4;
        let d = row * ow * 4;
        op[d..d + ow * 4].copy_from_slice(&sp[s..s + ow * 4]);
    }
    Some(out)
}

/// The laid-out box of the control the user named, from the document.
///
/// Matched on the `id` — what every frontend lowers a control's `Name` to, and
/// therefore the name a user types. Exact and case-insensitive first, then an
/// UNAMBIGUOUS substring, because exact names are often unusable: an SDL
/// surface is named after the window title, so it can be
/// `vybe sdl adapter - signal monitor_surface` and the user types
/// `--capture-control surface`.
fn control_rect(name: &str) -> Result<widgets::LayoutRect, String> {
    let controls = crate::gui_document::controls();
    let needle = name.to_lowercase();

    let exact = controls
        .iter()
        .find(|c| c.id.eq_ignore_ascii_case(name))
        .and_then(|c| c.rect);
    if let Some(rect) = exact {
        return Ok(rect);
    }

    let hits: Vec<&crate::gui_document::DomControl> = controls
        .iter()
        .filter(|c| c.id.to_lowercase().contains(&needle))
        .collect();
    if let [only] = hits.as_slice() {
        if let Some(rect) = only.rect {
            return Ok(rect);
        }
    }

    // ⚠ A control with no rect is NOT missing — it is present and never laid
    // out, which is a different bug and worth saying so.
    if let Some(found) = controls.iter().find(|c| c.id.eq_ignore_ascii_case(name)) {
        return Err(format!(
            "control `{name}` is in the document but has no laid-out box (connected: {})",
            found.connected
        ));
    }

    let mut names: Vec<&str> = controls.iter().map(|c| c.id.as_str()).collect();
    names.sort_unstable();
    names.dedup();
    Err(format!(
        "no control named `{name}` (have: {})",
        names.join(", ")
    ))
}

/// Render one frame offscreen and write it to `path` as a PNG.
///
/// `control` crops to that control's rect, matched on its `id`. `None` captures
/// the whole frame. Returns the written size.
///
/// Builds its own `FontSystem`, so it needs no live window — that is what lets
/// the step debugger and the headless flag share it.
pub fn capture_to_png(
    path: &str,
    control: Option<&str>,
    scale: f32,
) -> Result<(u32, u32), String> {
    // The frame IS the document's viewport: a form's `Width`/`Height` are CSS
    // on the body and land there.
    // The failure says WHAT THE ENGINE HOLDS, not just that it holds nothing
    // useful. "no live document" is true of an empty document and of a document
    // the guest built into an engine that is not the live one, and those are
    // completely different bugs — the first is a program that drew nothing, the
    // second is a swap that only half happened.
    let (w, h) = crate::gui_document::viewport().ok_or_else(|| {
        format!(
            "no live document to capture — engine `{}` holds: {}",
            vybe_platform_web::engine_select::live()
                .map(|e| e.as_str())
                .unwrap_or("none"),
            crate::gui_document::engine_html()
        )
    })?;
    let (w, h) = (w.max(1) as f32, h.max(1) as f32);
    let pw = (w * scale).round().max(1.0) as u32;
    let ph = (h * scale).round().max(1.0) as u32;
    let mut pixmap = Pixmap::new(pw, ph).ok_or_else(|| format!("bad frame size {pw}x{ph}"))?;

    let mut font_system = FontSystem::new();
    let mut swash_cache = SwashCache::new();
    render_into(&mut pixmap, scale);

    let shot = match control {
        None => pixmap,
        Some(name) => {
            let rect = control_rect(name)?;
            crop(
                &pixmap,
                rect.x * scale,
                rect.y * scale,
                rect.w * scale,
                rect.h * scale,
            )
            .ok_or_else(|| format!("control `{name}` has an empty rect"))?
        }
    };

    let png = shot.encode_png().map_err(|e| format!("encode: {e}"))?;
    std::fs::write(path, png).map_err(|e| format!("write {path}: {e}"))?;
    Ok((shot.width(), shot.height()))
}
