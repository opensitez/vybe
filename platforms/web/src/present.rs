//! Getting a frame out of whichever engine is live.
//!
//! # In-process frame capture
//!
//! `render` hands the engine a `tiny_skia::Pixmap` and says "paint here". That
//! works for WebCore because it rasterises with tiny-skia in this
//! process. **Chrome cannot do it.** An out-of-process engine owns its own
//! compositor: you give it a native window handle, or it hands YOU a buffer it
//! produced. A host-owned pixmap in the signature is this build's convenience
//! baked into the seam, exactly as `Op2D::SetFillStyle(u8, u8, u8, u8)` baked in
//! a pre-parsed colour.
//!
//! Windowed Webcore presentation is owned by `webcore::embedded_window`; this
//! in-process pixmap path is also used for headless capture. An out-of-process
//! browser would provide pixels through its own capture API.
//!
//! What this DOES fix is `vybex` reaching around the intermediary. It called
//! the old native toolkit directly, so the engine could swap and the
//! renderer could not. Everything above still asks `platforms/web`, which is
//! the arrangement that stays true when the answer moves.

use tiny_skia::Pixmap;

use crate::engine::DocumentId;
use crate::engine_select::{self, Engine};

/// Did the guest build a UI in this document?
///
/// `html::active_document()` CREATES a document on first call, so its existence
/// proves nothing — a program that never touched `web:*` still has one. Having
/// content is the test, and it is the same question `render` asks before it
/// paints.
///
/// One question, asked of whichever engine is live — and asked as a SELECTOR,
/// which is the vocabulary both already share.
///
/// It used to be two engine-specific questions. The toolkit's was
/// `form().control_count() > 0`, and every document answers at least `1` to
/// that: `Document::new` appends `<html>` as a control, so an empty document
/// claimed to have a UI. A parsed document always has the `html`/`head`/`body`
/// skeleton too, so the same is true there — what counts under either engine is
/// whether anything was put IN the body.
pub fn has_content(document: DocumentId) -> bool {
    matches!(
        crate::engine::apply(
            document,
            crate::engine::DomOp::QuerySelector("body > *".into()),
        ),
        crate::engine::DomValue::Node(_)
    )
}

/// Paint the live document into `pixmap`. Answers whether anything was drawn.
///
/// See the module note: the pixmap in this signature is the in-process
/// shortcut, not the destination.
pub fn render(document: DocumentId, pixmap: &mut Pixmap, scale: f32) -> bool {
    if !has_content(document) {
        return false;
    }
    match engine_select::live() {
        #[cfg(feature = "engine-webcore")]
        Some(Engine::WebCore) => render_webcore(document, pixmap, scale),
        _ => false,
    }
}

#[cfg(feature = "engine-webcore")]
fn render_webcore(document: DocumentId, pixmap: &mut Pixmap, scale: f32) -> bool {
    // webcore's `Renderer` is STATEFUL across frames — it carries the
    // compositor layer tree, the tile cache and the display list. Rebuilding it
    // per frame would not be a slow correct renderer, it would be one that
    // never gets to use any of that, so it is kept for the life of the thread.
    thread_local! {
        static RENDERER: std::cell::RefCell<webcore::renderer::Renderer> =
            std::cell::RefCell::new(webcore::renderer::Renderer::new());
    }
    RENDERER.with(|renderer| {
        let mut renderer = renderer.borrow_mut();
        crate::engine_webcore::with_document_resources(document, |doc, resources_dirty| {
            // **Layout, THEN paint** — the order webcore's own window runs
            // (`webcore::main`), and through the renderer's engine so its
            // caches see the same generation the paint reads. Painting without
            // it draws a tree that has no geometry: every box at the origin
            // with zero size, which comes out as a blank frame with no error.
            //
            // The viewport is the frame, in CSS pixels: the pixmap is in DEVICE
            // pixels, so a HiDPI frame is `scale` times larger and laying out
            // against it would make the page think the window is twice as wide.
            let width = pixmap.width() as f32 / scale;
            let height = pixmap.height() as f32 / scale;
            renderer.set_scale(scale);
            // ⛔ **Lay out EVERY frame. Do not gate this on the dirty flags.**
            //
            // I gated it on `layout_dirty || has_dirty_descendant ||
            // has_dirty_layout_descendant` to stop a full layout per 60Hz tick,
            // and it broke the page on interaction: clicking blanked the app
            // bar, the status text and the button. A click changes state
            // through paths that do not raise those flags — `handle_form_click`
            // is a free function over `&mut WebCore` with no document to mark —
            // so the frame after a click painted a tree that had never been
            // laid out at the current viewport.
            //
            // The flags are the engine's own INTERNAL pruning signal, read
            // partway down `layout_box`, where the tree walk has already
            // established what is being asked. They are not an "is this
            // document up to date" answer for a caller outside layout, and
            // using them as one is what made the difference invisible until
            // something interacted.
            //
            // The per-frame cost is real and the fix belongs in the engine: a
            // document that knows it is clean should make `layout()` cheap,
            // rather than a caller guessing from the outside.
            //
            // ONE borrow of `layout_engine()`: it re-seeds the engine from the
            // renderer on every call — including `viewport_h` — so setting the
            // height through one call and laying out through the next throws
            // the height away, and `html { height: 100% }` then has nothing to
            // resolve against.
            let engine = renderer.layout_engine();
            engine.viewport_h = height;
            engine.layout(doc, width);
            if *resources_dirty {
                webcore::restart_async_image_fetches(doc);
                *resources_dirty = false;
            }
            doc.poll_pending_images_budgeted(64, std::time::Duration::ZERO);
            renderer.render(doc, pixmap, scale);
            true
        })
        .unwrap_or(false)
    })
}
