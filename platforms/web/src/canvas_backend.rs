//! The seam between `web:canvas` (the API) and whatever actually paints.
//!
//! `CanvasRenderingContext2D` is a web-platform interface, so it is declared
//! here; the pixels belong to an engine. `widgets` is that engine today
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
    /// `transform(a, b, c, d, e, f)` — MULTIPLY the current matrix by this one.
    ///
    /// `translate`/`scale`/`rotate` are three special cases of it, and a caller
    /// holding a matrix (every 2D scene graph does) can express it no other
    /// way. `setTransform` is this preceded by [`Op2D::ResetTransform`] —
    /// composing versus replacing is the whole difference between the two, so
    /// they are not one op with a flag.
    Transform(f32, f32, f32, f32, f32, f32),
    /// `resetTransform()` — back to the identity matrix.
    ResetTransform,
    /// `miterLimit` — how far a mitred join may extend before it is bevelled.
    /// Without it a sharp corner spikes arbitrarily far from the path.
    SetMiterLimit(f32),
    /// `lineDashOffset` — where in the dash pattern a line starts. Animating it
    /// is what marching ants is.
    SetLineDashOffset(f32),
    /// `textAlign` / `textBaseline` — which end of the text `x` names, and
    /// which line of it `y` names. Carried as the spec's keyword so the painter
    /// owns the parse and a second spelling cannot appear here.
    SetTextAlign(String),
    SetTextBaseline(String),

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
    /// `ellipse(x, y, rx, ry, …)` — an axis-aligned ellipse. The engine has
    /// implemented it since the trait was written; only the wire format was
    /// missing, so no page could reach it.
    Ellipse(f32, f32, f32, f32),
    Fill,
    Stroke,
    Clip,

    // ── shapes, text, images ─────────────────────────────────────────────
    FillRect(f32, f32, f32, f32),
    StrokeRect(f32, f32, f32, f32),
    ClearRect(f32, f32, f32, f32),
    FillText(String, f32, f32),
    StrokeText(String, f32, f32),
    /// `putImageData(imagedata, dx, dy)` — a RAW pixel write.
    ///
    /// The spec route for handing a computed frame to a canvas, and the only
    /// one: `drawImage` takes an image SOURCE (an element, an `ImageBitmap`),
    /// never a byte array, so a software renderer that has pixels and no
    /// element has exactly this door.
    ///
    /// Unaffected by the transform, the clip, `globalAlpha` and the
    /// compositing mode (HTML §4.12.5) — the backend must write, not paint.
    /// There is no `dw`/`dh`: `putImageData` does not scale.
    PutImageData {
        pixels: Vec<u8>,
        width: u32,
        height: u32,
        dx: f32,
        dy: f32,
    },
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

    // ── Values that stay STRINGS across the seam ─────────────────────────
    //
    // `fillStyle`, `font`, `filter` and the rest are CSS values in the IDL, and
    // a page assigns them as written. They cross as the author's text and are
    // parsed by the ENGINE, with the engine's own CSS parser — the same one the
    // page's stylesheet goes through.
    //
    // The pre-parsed forms above (`SetFillStyle(u8, u8, u8, u8)`,
    // `SetFont { family, size, .. }`) are kept, not replaced: .NET's
    // `System.Drawing` and SDL emit them today and hold a colour or a font
    // struct rather than a string. Two ways in, one meaning.
    /// `fillStyle = "..."` — any CSS `<color>`.
    SetFillStyleCss(String),
    /// `strokeStyle = "..."`
    SetStrokeStyleCss(String),
    /// `font = "..."` — a CSS `font` shorthand. Kept verbatim by the engine so
    /// the attribute reads back what was set, which §4.12.5 requires.
    SetFontCss(String),
    /// `filter = "..."` — a CSS `<filter-value-list>`.
    SetFilter(String),
    /// `globalCompositeOperation = "..."` — one of the 26 blend keywords.
    SetGlobalCompositeOperation(String),
    /// `imageSmoothingQuality = "low" | "medium" | "high"`
    SetImageSmoothingQuality(String),

    // ── Shadows ──────────────────────────────────────────────────────────
    /// `shadowColor = "..."`
    SetShadowColor(String),
    /// `shadowBlur` — TWICE the Gaussian standard deviation, unlike CSS
    /// `drop-shadow()`, which names the deviation itself.
    SetShadowBlur(f32),
    SetShadowOffsetX(f32),
    SetShadowOffsetY(f32),

    // ── Text style ───────────────────────────────────────────────────────
    SetDirection(String),
    SetLetterSpacing(String),
    SetWordSpacing(String),
    SetFontKerning(String),
    SetFontStretch(String),
    SetFontVariantCaps(String),
    SetTextRendering(String),
    SetLang(String),

    // ── Paths ────────────────────────────────────────────────────────────
    /// `arcTo(x1, y1, x2, y2, radius)` — the corner-rounding primitive, and
    /// the one `roundRect` is defined in terms of.
    ArcTo(f32, f32, f32, f32, f32),
    /// `roundRect(x, y, w, h, radii)` — four corner radii, already expanded
    /// from whatever short form the page wrote.
    RoundRect {
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        radii: [f32; 4],
    },
    /// `ellipse(x, y, rx, ry, rotation, start, end, ccw)` — the IDL's full
    /// eight arguments. [`Op2D::Ellipse`] is the four-argument form that
    /// predates it and drops rotation, both angles and the direction.
    EllipseFull {
        x: f32,
        y: f32,
        rx: f32,
        ry: f32,
        rotation: f32,
        start: f32,
        end: f32,
        ccw: bool,
    },
    /// `fill(fillRule)` / `clip(fillRule)` — `"nonzero"` or `"evenodd"`.
    /// [`Op2D::Fill`] and [`Op2D::Clip`] are the default-rule forms.
    FillWithRule(String),
    ClipWithRule(String),

    // ── Text with a maximum width ────────────────────────────────────────
    /// `fillText(text, x, y, maxWidth)` — the string is CONDENSED to fit
    /// rather than clipped, per §4.12.5.
    FillTextMaxWidth(String, f32, f32, f32),
    StrokeTextMaxWidth(String, f32, f32, f32),

    // ── Gradients and patterns ───────────────────────────────────────────
    //
    // A gradient crosses whole rather than by handle. `createLinearGradient`
    // answers an object the page holds and adds stops to; assigning it to
    // `fillStyle` sends what it has become. No handle means no registry, no
    // lifetime to get wrong, and nothing to leak when a page drops one.
    /// `fillStyle = <CanvasGradient>`
    SetFillGradient(GradientDef),
    SetStrokeGradient(GradientDef),
    /// `fillStyle = <CanvasPattern>`
    SetFillPattern(PatternDef),
    SetStrokePattern(PatternDef),

    // ── The rest ─────────────────────────────────────────────────────────
    /// `reset()` — clears the bitmap, the state and the path.
    Reset,
    /// `putImageData(data, dx, dy, dirtyX, dirtyY, dirtyW, dirtyH)`
    PutImageDataDirty {
        pixels: Vec<u8>,
        width: u32,
        height: u32,
        dx: f32,
        dy: f32,
        dirty_x: f32,
        dirty_y: f32,
        dirty_w: f32,
        dirty_h: f32,
    },
    /// `drawFocusIfNeeded(element)` — draws a focus ring when the element it
    /// names has focus. The seam carries the decision, not the element.
    DrawFocusIfNeeded(bool),

    // ── Path2D ───────────────────────────────────────────────────────────
    //
    // `fill(path)`, `stroke(path)` and `clip(path)` are the IDL's overloads
    // that take an explicit path instead of the context's current one. They do
    // NOT disturb the current path, which is the difference that makes a
    // `Path2D` worth having.
    /// `fill(path, fillRule)`
    FillPath(PathDef, String),
    /// `stroke(path)`
    StrokePath(PathDef),
    /// `clip(path, fillRule)`
    ClipPath(PathDef, String),
    /// `ctx.addPath` has no IDL form; this is `Path2D.addPath` applied to the
    /// CONTEXT's current path, which is how a page folds a prebuilt shape into
    /// what it is drawing.
    AppendPath(PathDef),
}

/// A `Path2D`, as the page built it.
///
/// **Carried whole, like a gradient, rather than by handle.** A `Path2D` is
/// built once and used many times — that is its whole reason to exist, since
/// the context's own current path is consumed by `fill()` and thrown away by
/// `clip()`. Sending the operations when it is USED means no registry of live
/// paths, nothing to leak when a page drops one, and no way for the engine's
/// copy to fall out of step with the page's.
#[derive(Clone, Debug, Default)]
pub struct PathDef {
    pub ops: Vec<PathOp2D>,
}

/// One segment of a [`PathDef`].
///
/// Mirrors the engine's own path operations exactly. A separate enum because a
/// seam type cannot name an engine type — and because a missing arm in a
/// backend's conversion is then a compile error rather than a segment that
/// silently does not draw.
#[derive(Clone, Copy, Debug)]
pub enum PathOp2D {
    ClosePath,
    MoveTo(f32, f32),
    LineTo(f32, f32),
    QuadraticCurveTo { cx: f32, cy: f32, x: f32, y: f32 },
    BezierCurveTo { cx1: f32, cy1: f32, cx2: f32, cy2: f32, x: f32, y: f32 },
    ArcTo { x1: f32, y1: f32, x2: f32, y2: f32, radius: f32 },
    Rect { x: f32, y: f32, w: f32, h: f32 },
    RoundRect { x: f32, y: f32, w: f32, h: f32, radii: [f32; 4] },
    Arc { x: f32, y: f32, r: f32, start: f32, end: f32, ccw: bool },
    Ellipse { x: f32, y: f32, rx: f32, ry: f32, rotation: f32, start: f32, end: f32, ccw: bool },
}

/// A gradient, as the page built it.
#[derive(Clone, Debug)]
pub struct GradientDef {
    /// `linear` = (x0, y0, x1, y1); `radial` = (x0, y0, r0, x1, y1, r1);
    /// `conic` = (angle, x, y).
    pub kind: GradientKind,
    /// `addColorStop(offset, color)` — offset in 0..1 and any CSS `<color>`,
    /// in the order the page added them.
    pub stops: Vec<(f32, String)>,
}

#[derive(Clone, Debug)]
pub enum GradientKind {
    Linear {
        x0: f32,
        y0: f32,
        x1: f32,
        y1: f32,
    },
    Radial {
        x0: f32,
        y0: f32,
        r0: f32,
        x1: f32,
        y1: f32,
        r1: f32,
    },
    Conic {
        angle: f32,
        x: f32,
        y: f32,
    },
}

/// A pattern, as the page built it: `createPattern(image, repetition)`.
#[derive(Clone, Debug)]
pub struct PatternDef {
    pub pixels: Vec<u8>,
    pub width: u32,
    pub height: u32,
    /// `"repeat"`, `"repeat-x"`, `"repeat-y"` or `"no-repeat"`.
    pub repetition: String,
}

/// Something a page ASKS a canvas, rather than tells it.
///
/// **Why this is a second enum and not more [`Op2D`].** An op is delivered and
/// forgotten — [`apply`] returns nothing, which is what lets a paint call on a
/// canvas with no renderer be invisible rather than an error. A question has to
/// come back with an answer, and an answer that goes missing cannot be silent:
/// a caller laying out against a measurement it did not really get would be
/// wrong in a way nothing could see.
///
/// Every one of these is a question about state the PREVIOUS ops established,
/// so a backend has to answer from the same retained context it paints into.
#[derive(Clone, Debug)]
pub enum Query2D {
    /// `measureText(text)` — the full `TextMetrics`, not just the width.
    MeasureText(String),
    /// `getImageData(sx, sy, sw, sh)` — STRAIGHT (un-premultiplied) RGBA.
    ///
    /// Copies the requested rectangle out of the surface, so a caller reading
    /// a whole canvas every frame is copying a whole canvas every frame. That
    /// is inherent to the spec's design, not to this seam — but it is worth
    /// knowing before putting one in a render loop.
    GetImageData {
        sx: i32,
        sy: i32,
        sw: u32,
        sh: u32,
    },
    /// `isPointInPath(x, y, fillRule)` — the point is in the space the page's
    /// own transform maps into, and is mapped back through it.
    IsPointInPath {
        x: f32,
        y: f32,
        rule: String,
    },
    /// `isPointInStroke(x, y)`
    IsPointInStroke {
        x: f32,
        y: f32,
    },
    /// `getTransform()` — the page's own matrix, `[a, b, c, d, e, f]`.
    GetTransform,
    /// `getLineDash()` — after the spec's normalisation, so an odd-length list
    /// reads back doubled.
    GetLineDash,
    /// `isContextLost()`
    IsContextLost,
    /// `canvas.toDataURL(type, quality)`
    ToDataUrl {
        mime: String,
        quality: Option<f32>,
    },
    /// `canvas.toBlob(callback, type, quality)` — the encoded bytes.
    ToBlob {
        mime: String,
        quality: Option<f32>,
    },
    /// `getContextAttributes()` — the settings the context was created with.
    GetContextAttributes,
    /// `isPointInPath(path, x, y, fillRule)` — the explicit-path overload.
    /// Asks about the PATH given, leaving the context's current path alone.
    IsPointInPathOf {
        path: PathDef,
        x: f32,
        y: f32,
        rule: String,
    },
    /// `isPointInStroke(path, x, y)`
    IsPointInStrokeOf { path: PathDef, x: f32, y: f32 },
    /// `font` / `fillStyle` / `strokeStyle` / `filter` and the other string
    /// attributes, read back. \S4.12.5 requires these to serialize, and for a
    /// colour or a font the engine keeps the text it was given.
    GetStringAttribute(StringAttribute),
}

/// Which string attribute [`Query2D::GetStringAttribute`] is asking about.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StringAttribute {
    Font,
    FillStyle,
    StrokeStyle,
    Filter,
    GlobalCompositeOperation,
    ImageSmoothingQuality,
    ShadowColor,
    Direction,
    LetterSpacing,
    WordSpacing,
    FontKerning,
    FontStretch,
    FontVariantCaps,
    TextRendering,
    Lang,
    TextAlign,
    TextBaseline,
    LineCap,
    LineJoin,
}

/// What a [`Query2D`] answers.
///
/// `Absent` is not an error and not a zero — it is "this target names no
/// surface". A caller has to be able to tell that from a real answer, which is
/// the whole reason questions do not go through [`apply`].
#[derive(Clone, Debug)]
pub enum Query2DValue {
    Absent,
    Bool(bool),
    Text(String),
    /// `[a, b, c, d, e, f]`
    Matrix([f32; 6]),
    Floats(Vec<f32>),
    Bytes(Vec<u8>),
    /// Straight RGBA plus the dimensions it is laid out in.
    Pixels {
        data: Vec<u8>,
        width: u32,
        height: u32,
    },
    Metrics(TextMetrics2D),
    /// `getContextAttributes()` — `(alpha, desynchronized, colorSpace,
    /// colorType, willReadFrequently)`.
    ContextAttributes {
        alpha: bool,
        desynchronized: bool,
        color_space: String,
        color_type: String,
        will_read_frequently: bool,
    },
}

/// `TextMetrics` — HTML §4.12.5.
///
/// All twelve members. The seam used to carry `width` alone, and the other
/// eleven were computed by the engine and dropped on the way out.
#[derive(Clone, Copy, Debug, Default)]
pub struct TextMetrics2D {
    pub width: f32,
    pub actual_bounding_box_left: f32,
    pub actual_bounding_box_right: f32,
    pub actual_bounding_box_ascent: f32,
    pub actual_bounding_box_descent: f32,
    pub font_bounding_box_ascent: f32,
    pub font_bounding_box_descent: f32,
    pub em_height_ascent: f32,
    pub em_height_descent: f32,
    pub hanging_baseline: f32,
    pub alphabetic_baseline: f32,
    pub ideographic_baseline: f32,
}

/// What a painting engine must provide for `web:canvas` to work.
pub trait CanvasBackend: Send + Sync {
    /// Apply `op` to the drawing surface named `target`. Creating storage for
    /// an unknown target is the backend's business — the API only ever names
    /// one, the same way a page names an element id.
    fn apply(&self, target: &str, op: Op2D);

    /// Ensure a surface exists for `target` (`getContext`'s side effect).
    fn ensure(&self, target: &str);

    /// Answer a question about the surface `target` names.
    ///
    /// **The half of the API that asks.** [`apply`] is fire-and-forget, which
    /// is right for painting — a draw call on a canvas with no renderer should
    /// be invisible, not an error. It is wrong for a question: a measurement
    /// that goes nowhere has to return SOMETHING, and any value it invents is
    /// wrong rather than absent, so an adapter would lay out against it and
    /// never know. [`Query2DValue::Absent`] is how a backend says "no surface",
    /// and a caller that ignores it has made that choice visibly.
    ///
    /// This used to be a single `measure_text` returning `Option<f32>` —
    /// `measureText` was the only thing a page could ask, and it got back one
    /// of `TextMetrics`' twelve members.
    fn query(&self, target: &str, q: Query2D) -> Query2DValue;

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

/// Ask one question, answering [`Query2DValue::Absent`] when no engine is
/// installed. A question with nobody to answer it is absent, never a default:
/// the whole point of the type is that "there is no surface" and "the answer
/// happens to be zero" cannot be confused.
pub fn query(target: &str, q: Query2D) -> Query2DValue {
    match backend() {
        Some(b) => b.query(target, q),
        None => Query2DValue::Absent,
    }
}
