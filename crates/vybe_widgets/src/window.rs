//! WHATWG HTML §7 — browsing contexts and the `Window` interface.
//!
//! `open(target, features)` creates a browsing context AND its initial
//! `about:blank` document, then hands back the `Window`. That is the spec's
//! own bootstrap: a page builds a new window's contents by opening it and
//! calling `createElement` **on that window's document**. Nothing here
//! invents a `createWindow`.
//!
//! A window's size is NOT stored twice. `innerWidth`/`innerHeight` are read
//! back off the document's viewport — the body's containing block — because
//! that is the same measurement, and keeping two copies is how they drift.
//!
//! One honest difference from a browser: there, `open` is a method on an
//! existing `Window`, because the user agent already made a tab. A standalone
//! toolkit has no tab — it may legitimately have no window at all — so the
//! first `open` is a free function. Everything after is standard.
//!
//! A VB form, a Pascal form and a Flutter window are all "a window": the
//! form's controls live in its document, while Flutter paints into a canvas
//! in its document — same context, different content.

use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};

use crate::dom::{self, DocumentId};

/// A browsing context handle — what `open()` returns.
pub type WindowId = u64;

/// One browsing context: its document, its position, its lifecycle.
pub struct BrowsingContext {
    pub id: WindowId,
    /// `window.document`. The context holds the handle; the tree is the
    /// document's.
    pub document_id: DocumentId,
    /// `window.name`, the `target` of `open()`.
    pub name: String,
    /// `screenX` / `screenY`. Unlike the size, this has no document
    /// counterpart — a page cannot see where its window sits on screen.
    pub screen_x: f64,
    pub screen_y: f64,
    /// `window.closed`.
    pub closed: bool,
}

#[derive(Default)]
struct Contexts {
    windows: HashMap<WindowId, BrowsingContext>,
    order: Vec<WindowId>,
    next_id: WindowId,
}

fn contexts() -> &'static Mutex<Contexts> {
    static CTX: OnceLock<Mutex<Contexts>> = OnceLock::new();
    CTX.get_or_init(|| Mutex::new(Contexts::default()))
}

/// Parse a `windowFeatures` string — `"width=800,height=600,left=10"`. The
/// spec's own comma-separated `name=value` form, ignoring what it does not
/// know, exactly as a user agent does.
fn parse_features(features: &str) -> HashMap<String, f64> {
    let mut out = HashMap::new();
    for part in features.split(',') {
        let Some((k, v)) = part.split_once('=') else {
            continue;
        };
        if let Ok(n) = v.trim().parse::<f64>() {
            out.insert(k.trim().to_ascii_lowercase(), n);
        }
    }
    out
}

/// `window.open(url, target, features)` → a new browsing context with a fresh
/// document. `url` is accepted and ignored: there is no navigation here, so
/// every window opens the spec's initial `about:blank`.
pub fn open(target: &str, features: &str) -> WindowId {
    let f = parse_features(features);
    let document_id = dom::new_document(target);
    let width = f.get("width").copied().unwrap_or(800.0);
    let height = f.get("height").copied().unwrap_or(600.0);
    // The viewport IS the window size — set once, read back from there.
    dom::with_document(document_id, |d| d.set_viewport(width as f32, height as f32));

    let mut ctx = contexts().lock().unwrap();
    ctx.next_id += 1;
    let id = ctx.next_id;
    ctx.windows.insert(
        id,
        BrowsingContext {
            id,
            document_id,
            name: target.to_string(),
            screen_x: f.get("left").copied().unwrap_or(0.0),
            screen_y: f.get("top").copied().unwrap_or(0.0),
            closed: false,
        },
    );
    ctx.order.push(id);
    id
}

/// Give an EXISTING document its top-level browsing context — the tab the user
/// agent already made.
///
/// **This is the half of HTML §7 that was missing.** `open()` creates a context
/// AND a document, which is right for `window.open()`, and it was the ONLY way a
/// context came into being — so the ambient document, the one every program
/// actually runs in, had none. In a browser that cannot happen: the initial
/// document always sits in a top-level traversable, because the tab exists
/// before a byte of script does. Everything downstream followed from the gap —
/// `Document.defaultView` answered null for every program, and `window.close()`
/// was unreachable for the main window because nothing could name it.
///
/// Idempotent, and it must be: `active_document()` is called on nearly every
/// DOM operation, and a second context over one document would be a second tab
/// showing the same page.
///
/// Marked `#[must_use]` for the id, but the id is genuinely ignorable at the
/// bootstrap call site — what matters there is that the context now EXISTS.
pub fn adopt(document_id: DocumentId, name: &str) -> WindowId {
    let mut ctx = contexts().lock().unwrap();
    // `.map` ends the immutable borrow before the insert below.
    if let Some(existing) = ctx
        .windows
        .values()
        .find(|w| w.document_id == document_id)
        .map(|w| w.id)
    {
        return existing;
    }
    ctx.next_id += 1;
    let id = ctx.next_id;
    ctx.windows.insert(
        id,
        BrowsingContext {
            id,
            document_id,
            name: name.to_string(),
            // A top-level context starts at the origin. `open()` takes
            // `left`/`top` from its features string; nobody passes features to
            // the tab they were already given.
            screen_x: 0.0,
            screen_y: 0.0,
            closed: false,
        },
    );
    ctx.order.push(id);
    id
}

/// `window.screen` (CSSOM View) — the DISPLAY, not the window.
///
/// A toolkit has no display of its own to measure, and the SHELL does — winit
/// knows the monitor, this crate does not. So the shell sets it and everything
/// else reads it, rather than each caller guessing.
///
/// ⛔ Unset, this answers the context's own VIEWPORT, which makes "centre on the
/// screen" resolve to "already centred". That is deliberate: the retired
/// `vybe:gui::__form_center_to_screen` hardcoded **1920x1080** with a comment
/// admitting it was a guess, and a wrong number that lands a window off-screen
/// is worse than a true statement that nothing is known. A real engine supplies
/// the real display and the arithmetic starts meaning something with no code
/// change here.
fn screen_slot() -> &'static Mutex<Option<(f64, f64)>> {
    static SCREEN: OnceLock<Mutex<Option<(f64, f64)>>> = OnceLock::new();
    SCREEN.get_or_init(|| Mutex::new(None))
}

/// Tell the registry how big the real display is. For the shell that owns a
/// window system; a browser engine would answer from the platform instead.
pub fn set_screen(width: f64, height: f64) {
    *screen_slot().lock().unwrap() = Some((width, height));
}

/// `screen.width` / `screen.height`, falling back to this context's viewport.
pub fn screen(id: WindowId) -> (f64, f64) {
    match *screen_slot().lock().unwrap() {
        Some(size) => size,
        None => inner_size(id),
    }
}

/// `window.focus()` (HTML §7.2.2) — bring this context to the front.
///
/// Front-ness IS `order`: the list is the z-order every consumer already reads
/// through [`open_windows`], so focusing moves the id to the end rather than
/// storing a second `focused` flag that could disagree with it. One fact, one
/// place — the same rule `innerWidth` follows by reading the viewport instead
/// of keeping its own copy of the size.
pub fn focus(id: WindowId) {
    let mut ctx = contexts().lock().unwrap();
    if !ctx.windows.contains_key(&id) {
        return;
    }
    ctx.order.retain(|w| *w != id);
    ctx.order.push(id);
}

/// `Document.defaultView` — the window whose document this is.
///
/// The inverse of [`document`], and the standard way a script reaches its own
/// window when it holds a document rather than the global. Null for a document
/// with no context, which is the spec's own answer (a `DOMParser` document has
/// no browsing context and no `defaultView`).
///
/// A CLOSED context still answers, deliberately: `w.close(); w.document` is
/// valid, and a closed window is a live handle reporting `closed = true`.
pub fn default_view(document_id: DocumentId) -> Option<WindowId> {
    let ctx = contexts().lock().unwrap();
    ctx.windows
        .values()
        .find(|w| w.document_id == document_id)
        .map(|w| w.id)
}

pub fn with_window<T>(id: WindowId, f: impl FnOnce(&BrowsingContext) -> T) -> Option<T> {
    let ctx = contexts().lock().unwrap();
    ctx.windows.get(&id).map(f)
}

fn with_window_mut<T>(id: WindowId, f: impl FnOnce(&mut BrowsingContext) -> T) -> Option<T> {
    let mut ctx = contexts().lock().unwrap();
    ctx.windows.get_mut(&id).map(f)
}

/// `window.document`.
pub fn document(id: WindowId) -> Option<DocumentId> {
    with_window(id, |w| w.document_id)
}

/// `window.close()`.
pub fn close(id: WindowId) {
    with_window_mut(id, |w| w.closed = true);
}

/// `window.closed`. An unknown handle reads closed, as a stale reference does.
pub fn closed(id: WindowId) -> bool {
    with_window(id, |w| w.closed).unwrap_or(true)
}

/// `window.innerWidth` / `innerHeight` — the document's viewport, not a copy.
pub fn inner_size(id: WindowId) -> (f64, f64) {
    let Some(doc) = document(id) else {
        return (0.0, 0.0);
    };
    dom::with_document(doc, |d| {
        let r = d.viewport();
        (r.w as f64, r.h as f64)
    })
    .unwrap_or((0.0, 0.0))
}

/// `window.resizeTo(width, height)` — resizes the viewport the body fills.
pub fn resize_to(id: WindowId, width: f64, height: f64) {
    if let Some(doc) = document(id) {
        dom::with_document(doc, |d| d.set_viewport(width as f32, height as f32));
    }
}

/// `window.moveTo(x, y)`.
pub fn move_to(id: WindowId, x: f64, y: f64) {
    with_window_mut(id, |w| {
        w.screen_x = x;
        w.screen_y = y;
    });
}

pub fn screen_position(id: WindowId) -> (f64, f64) {
    with_window(id, |w| (w.screen_x, w.screen_y)).unwrap_or((0.0, 0.0))
}

pub fn name(id: WindowId) -> String {
    with_window(id, |w| w.name.clone()).unwrap_or_default()
}

/// Every open browsing context, in creation order — what a renderer walks to
/// know which windows exist.
pub fn open_windows() -> Vec<WindowId> {
    let ctx = contexts().lock().unwrap();
    ctx.order
        .iter()
        .copied()
        .filter(|id| ctx.windows.get(id).map(|w| !w.closed).unwrap_or(false))
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn open_gives_the_window_its_own_document() {
        let a = open("a", "width=300,height=200");
        let b = open("b", "");
        assert_ne!(document(a), document(b), "each window has its own tree");
        assert_eq!(inner_size(a), (300.0, 200.0), "features seed the viewport");
        assert_eq!(inner_size(b), (800.0, 600.0), "the spec-ish default");
    }

    #[test]
    fn resizing_moves_the_documents_viewport_not_a_copy() {
        let w = open("r", "width=400,height=300");
        resize_to(w, 640.0, 480.0);
        assert_eq!(inner_size(w), (640.0, 480.0));
        // Read it off the document directly — one measurement, not two.
        let seen = dom::with_document(document(w).unwrap(), |d| {
            let r = d.viewport();
            (r.w as f64, r.h as f64)
        });
        assert_eq!(seen, Some((640.0, 480.0)));
    }

    #[test]
    fn a_closed_window_leaves_the_open_list() {
        let w = open("c", "");
        assert!(open_windows().contains(&w));
        assert!(!closed(w));
        close(w);
        assert!(closed(w));
        assert!(!open_windows().contains(&w));
    }
}
