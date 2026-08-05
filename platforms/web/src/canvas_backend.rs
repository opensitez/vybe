//! The seam between `web:canvas` (the API) and whatever actually paints.
//!
//! `CanvasRenderingContext2D` is a web-platform interface, so it is declared
//! here; the pixels belong to an engine. `vybe_widgets` is that engine today
//! and a real browser engine could be tomorrow — neither is named by this
//! module. A host installs its painter with [`set_backend`] at startup and
//! the API surface never learns which one it got.
//!
//! This is the same shape as `web:timers` owning the wheel while the clock
//! comes from the runtime: the standard surface owns the contract, not the
//! machinery.

use std::sync::{Arc, OnceLock, RwLock};

/// One 2D drawing operation, in `CanvasRenderingContext2D` terms.
///
/// An enum rather than a 30-method trait so a backend implements ONE function
/// and can never silently miss an op — a missing arm is a compile error, and
/// adding an op fails every backend that hasn't handled it.
#[derive(Clone, Debug)]
pub enum Op2D {
    // ── state ────────────────────────────────────────────────────────────
    Save,
    Restore,
    /// `fillStyle = rgba(...)`
    SetFillStyle(u8, u8, u8, u8),
    /// `strokeStyle = rgba(...)`
    SetStrokeStyle(u8, u8, u8, u8),
    SetLineWidth(f32),
    /// `setLineDash([...])`; empty = solid
    SetLineDash(Vec<f32>),
    SetLineCap(String),
    SetLineJoin(String),
    SetGlobalAlpha(f32),
    /// `font = "<style> <weight> <size>px <family>"`, pre-parsed.
    SetFont {
        family: String,
        size: f32,
        bold: bool,
        italic: bool,
    },
    /// `imageSmoothingEnabled`
    SetImageSmoothing(bool),
    Translate(f32, f32),
    Scale(f32, f32),
    Rotate(f32),

    // ── paths ────────────────────────────────────────────────────────────
    BeginPath,
    ClosePath,
    MoveTo(f32, f32),
    LineTo(f32, f32),
    /// `arc(x, y, r, startAngle, endAngle, counterclockwise)`
    Arc(f32, f32, f32, f32, f32, bool),
    BezierCurveTo(f32, f32, f32, f32, f32, f32),
    QuadraticCurveTo(f32, f32, f32, f32),
    Rect(f32, f32, f32, f32),
    Fill,
    Stroke,
    Clip,

    // ── shapes, text, images ─────────────────────────────────────────────
    FillRect(f32, f32, f32, f32),
    StrokeRect(f32, f32, f32, f32),
    ClearRect(f32, f32, f32, f32),
    FillText(String, f32, f32),
    StrokeText(String, f32, f32),
    /// `drawImage` over dense RGBA pixels — `putImageData`'s cousin, and what
    /// a software renderer (SDL, Doom) hands over each frame.
    DrawImageRgba {
        pixels: Vec<u8>,
        width: u32,
        height: u32,
        dx: f32,
        dy: f32,
        dw: f32,
        dh: f32,
    },
    /// 8-bit paletted pixels expanded through a 256-entry RGB palette by the
    /// backend — the frame path of every palette-era game.
    DrawImagePaletted {
        indices: Vec<u8>,
        palette: Vec<u8>,
        width: u32,
        height: u32,
        dx: f32,
        dy: f32,
        dw: f32,
        dh: f32,
    },
}

/// What a painting engine must provide for `web:canvas` to work.
pub trait CanvasBackend: Send + Sync {
    /// Apply `op` to the drawing surface named `target`. Creating storage for
    /// an unknown target is the backend's business — the API only ever names
    /// one, the same way a page names an element id.
    fn apply(&self, target: &str, op: Op2D);

    /// Ensure a surface exists for `target` (`getContext`'s side effect).
    fn ensure(&self, target: &str);

    /// Drop everything drawn for `target`.
    fn clear_all(&self, target: &str);
}

fn slot() -> &'static RwLock<Option<Arc<dyn CanvasBackend>>> {
    static SLOT: OnceLock<RwLock<Option<Arc<dyn CanvasBackend>>>> = OnceLock::new();
    SLOT.get_or_init(|| RwLock::new(None))
}

/// Install the painting engine. Called once by whichever host owns a window.
pub fn set_backend(backend: Arc<dyn CanvasBackend>) {
    *slot().write().unwrap() = Some(backend);
}

pub fn backend() -> Option<Arc<dyn CanvasBackend>> {
    slot().read().unwrap().clone()
}

/// Apply one op, silently dropping it when no engine is installed — a page
/// with no renderer attached draws nothing; it does not fail.
pub fn apply(target: &str, op: Op2D) {
    if let Some(b) = backend() {
        b.apply(target, op);
    }
}
