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

use std::sync::{Arc, Mutex};

use vybe_platform_vybe::gui_state::GuiState;
use vybe_widgets::{FontSystem, PanelWidget, Pixmap, RenderContext, SwashCache, fill_background};

/// Paint the form and its overlay canvases into `pixmap`.
///
/// This is the whole frame: `FormApp::render` calls it for the live window and
/// the capture paths call it for an offscreen buffer.
pub fn render_into(
    gui: &Arc<Mutex<GuiState>>,
    pixmap: &mut Pixmap,
    font_system: &mut FontSystem,
    swash_cache: &mut SwashCache,
    scale: f32,
) {
    fill_background(pixmap, 240, 240, 240, 255);
    let mut g = match gui.lock() {
        Ok(g) => g,
        Err(poisoned) => poisoned.into_inner(),
    };
    let mut ctx = RenderContext {
        pixmap,
        font_system,
        swash_cache,
        scale,
    };
    // The window shows the DOCUMENT — `vybe_widgets` owns the tree and paints
    // it. A control is `document.createElement(tag)`, so the document's form
    // is the only one that has anything on it; `GuiState`'s own form is a
    // second, empty instance nothing writes to any more, and rendering that
    // one is why a form opened blank with every control present in the tree.
    //
    // Through `gui_document`, so this asks the same "is the document the live
    // tree" question the runner and the debugger ask. Reaching for
    // `html::active_document()` here instead painted an EMPTY document twice
    // over: once for a designer form that never opened one, and once for the
    // debugger's `capture`, which runs on the REPL thread where that
    // thread-local is a different document entirely.
    let document_painted =
        crate::gui_document::with_live(|document| document.render(&mut ctx)).is_some();
    if !document_painted {
        g.form.render(&mut ctx);
    }
    // The overlay compositing that used to run here is GONE, and it is the
    // reason this function could not be trusted: it painted
    // `GuiState.overlay_canvases` on top, which no window ever did. A canvas
    // drawing showed up in a capture and not on screen, so the screenshot said
    // the opposite of the truth.
    //
    // A `<canvas>` is an element in the document now and paints as part of it,
    // the same as any other control — so a capture and a window show the same
    // thing by construction rather than by agreement.
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

/// Render one frame offscreen and write it to `path` as a PNG.
///
/// `control` crops to that control's rect (case-insensitive, resolved the same
/// way host draw calls resolve names). `None` captures the whole form.
/// Returns the written size.
///
/// Builds its own `FontSystem`, so it needs no live window — that is what lets
/// the step debugger and the headless flag share it.
pub fn capture_to_png(
    gui: &Arc<Mutex<GuiState>>,
    path: &str,
    control: Option<&str>,
    scale: f32,
) -> Result<(u32, u32), String> {
    // The frame is the document's viewport when the document is the live tree
    // — a form's `Width`/`Height` are CSS on the body and land there, not in
    // `GuiState`, whose pair keeps its defaults.
    let (w, h) = {
        let g = gui.lock().map_err(|_| "gui state unavailable")?;
        let (w, h) = crate::gui_document::viewport().unwrap_or((g.width, g.height));
        (w.max(1) as f32, h.max(1) as f32)
    };
    let pw = (w * scale).round().max(1.0) as u32;
    let ph = (h * scale).round().max(1.0) as u32;
    let mut pixmap = Pixmap::new(pw, ph).ok_or_else(|| format!("bad frame size {pw}x{ph}"))?;

    // Lay the form out to the frame size, exactly as `on_init`/`on_resize` do
    // for a window. Without a window nothing has ever called those, so every
    // control still has a zero rect and the frame comes out blank — which is
    // what a `capture` at a debugger breakpoint hits.
    {
        let mut g = gui.lock().map_err(|_| "gui state unavailable")?;
        if g.form.rect().w < 1.0 || g.form.rect().h < 1.0 {
            g.form
                .set_rect(vybe_widgets::LayoutRect::new(0.0, 0.0, w, h));
        }
    }

    let mut font_system = FontSystem::new();
    let mut swash_cache = SwashCache::new();
    render_into(gui, &mut pixmap, &mut font_system, &mut swash_cache, scale);

    let shot = match control {
        None => pixmap,
        Some(name) => {
            let rect = {
                let g = gui.lock().map_err(|_| "gui state unavailable")?;
                let resolved = g.resolve_control_name(name);
                g.form
                    .get_control_rect(&resolved)
                    .or_else(|| g.form.get_control_rect(name))
                    // Exact names are often unusable: an SDL surface is named
                    // after the WINDOW TITLE, so it can be
                    // `vybe sdl adapter - signal monitor_surface`. Accept any
                    // unambiguous substring — `--capture-control surface`.
                    .or_else(|| {
                        let needle = name.to_lowercase();
                        let hits: Vec<&String> = g
                            .control_names
                            .iter()
                            .filter(|c| c.to_lowercase().contains(&needle))
                            .collect();
                        match hits.as_slice() {
                            // `control_names` stores the FOLDED name; the widget
                            // carries the original casing, so resolve before
                            // looking the rect up.
                            [only] => {
                                let canonical = g.resolve_control_name(only);
                                g.form
                                    .get_control_rect(&canonical)
                                    .or_else(|| g.form.get_control_rect(only))
                            }
                            _ => None,
                        }
                    })
                    // The overlay fallback that used to be here — "a surface
                    // has no rect of its own, so it covers the whole form" —
                    // is gone with the overlay map. An SDL surface IS a
                    // `<canvas>` element now, so it has a rect like anything
                    // else and the branch above finds it.
                    .ok_or_else(|| {
                        let mut names: Vec<&str> =
                            g.control_names.iter().map(|s| s.as_str()).collect();
                        names.sort_unstable();
                        names.dedup();
                        format!("no control named `{name}` (have: {})", names.join(", "))
                    })
            };
            let rect = rect?;
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
