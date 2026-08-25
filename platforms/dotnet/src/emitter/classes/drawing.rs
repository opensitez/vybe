//! `System.Drawing` value-shaped types: `Graphics`, `Pen`, `Brush`,
//! `SolidBrush`, `HatchBrush`, `LinearGradientBrush`.
//!
//! ## Architecture
//!
//! `Graphics` looks and acts exactly like .NET's `System.Drawing.Graphics`,
//! and every drawing call is an ADAPTER over **`web:canvas`** — the WHATWG
//! `CanvasRenderingContext2D` surface, resolving through the document. The
//! .NET API is translated here and nothing downstream knows GDI+ exists:
//! "DrawLine takes a Pen with a Color and a Width" becomes "set stroke style,
//! set line width, begin path, move to, line to, stroke". Same contract SDL
//! states on its side.
//!
//! Each `Graphics` method is a [`MethodTarget::Body`] sequence — a small
//! declarative slice of [`MethodOp`]s the builder compiles to bytecode. The
//! body reads `pen.color.r/g/b/a` and `pen.width` from the user's arguments
//! via `PushArgField` and forwards them to the spec'd ops.
//!
//! Where the two APIs describe a shape differently, the difference is
//! ARITHMETIC and it lives here: a .NET bounding box becomes a centre and
//! radii, degrees become radians. It belongs on this side of the seam — a
//! bespoke host function per .NET shape is not something a browser engine
//! could be asked to implement.
//!
//! `Control.CreateGraphics()` is `element.getContext("2d")` — see
//! `classes/control.rs::CONTROL_CREATE_GRAPHICS` and its twin in
//! `dispatch.rs`. The control IS the element, so the context binds to it and
//! the surface is the `<canvas>`'s own recording.
//!
//! ## Pen / Brush
//!
//! `Pen` and `SolidBrush` are real dotnet classes with arity-N
//! constructors. The user writes `New Pen(Color.Red, 5)` and the dotnet
//! ctor composes the object in bytecode via `dotnet.pen_new` — an Object with
//! `color` and `width` fields, built by `emit_value_type_new` and no host at
//! all. The Graphics method bodies read those fields directly.
//!
//! The classes also expose all the .NET property setters (`Pen.Color`,
//! `Pen.Width`, `Brush.Color`, etc.) via the standard property-setter
//! infrastructure, so user code can mutate a pen between draws and the
//! next DrawLine picks up the new state.
//!
//! ## Color
//!
//! `Color.Red`, `Color.Blue`, etc. are namespace constants — Objects
//! with `r/g/b/a` numeric fields (0-255). The named constants and
//! `Color.FromArgb` both produce the same shape, so the body sequences
//! that read `pen.color.r` work uniformly.

use super::{DotnetClass, DotnetMethod, MethodOp, MethodTarget};

/// Degrees → radians.
///
/// GDI+ states its angles in degrees; the canvas 2D context takes radians
/// (HTML §4.12.5). Both measure clockwise from the positive x-axis, so this
/// factor is the whole of the difference between them.
const DEG_TO_RAD: f64 = std::f64::consts::PI / 180.0;

// ─── Body templates ────────────────────────────────────────────────────────
//
// Each body is a static slice of MethodOp. The orchestrator pre-resolves
// every CallHost target to an import index in encounter order; the builder
// walks the slice and emits real bytecode. There's no runtime
// interpretation — `Body` lowers to identical bytecode to what hand-written
// ops would emit.

/// `Graphics.DrawLine(pen, x1, y1, x2, y2)`
///
/// Translation:
/// ```text
/// canvasSetStrokeColor(this, pen.color.r, .g, .b, .a)
/// canvasSetLineWidth(this, pen.width)
/// canvasBeginPath(this)
/// canvasMoveTo(this, x1, y1)
/// canvasLineTo(this, x2, y2)
/// canvasStroke(this)
/// ```
const GRAPHICS_DRAW_LINE: &[MethodOp] = &[
    // Apply pen state.
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "color"), // pen.color (object)
    // We need r/g/b/a as separate args, not the color object. The
    // canvas host fn expects 5 args: (handle, r, g, b, a). So we
    // extract the four channels via separate field reads on `pen`.
    //
    // Drop the color object we just pushed (we don't need it as a
    // single arg) and push the four channels individually instead.
    MethodOp::Drop,
    MethodOp::PushThis,
    // pen.color.r — there's no PushArgFieldField op, so we
    // synthesize: push pen.color, then access .r via SetField's
    // mirror... actually struct_get is what we need but our DSL
    // only provides PushArgField (one level deep).
    //
    // Workaround: extend the DSL OR use a helper. The cleanest fix
    // is to add `PushArgFieldField(arg_idx, f1, f2)`. Doing that.
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "moveTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "lineTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "stroke",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.DrawLine(pen, Point p1, Point p2)` — GDI+'s two-`Point` overload.
///
/// The canvas has one `moveTo`/`lineTo` shape and it takes numbers, so the only
/// difference from the four-coordinate body is where the numbers come FROM:
/// `p1.x` instead of arg 2. `pointNew` stores lowercase `x`/`y`, which is what
/// `PushArgField` reads — the property spelling `Point.X` is a separate axis.
///
/// Without this entry the call landed on the five-argument body and its
/// `PushArg(4)` ran off a four-slot frame, which is a PANIC rather than a miss.
const GRAPHICS_DRAW_LINE_POINTS: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(2, "x"),
    MethodOp::PushArgField(2, "y"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "moveTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(3, "x"),
    MethodOp::PushArgField(3, "y"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "lineTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "stroke",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.DrawRectangle(pen, x, y, w, h)`
const GRAPHICS_DRAW_RECTANGLE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "strokeRect",
        argc: 5,
    },
    MethodOp::Return,
];

/// `Graphics.FillRectangle(brush, x, y, w, h)`
const GRAPHICS_FILL_RECTANGLE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setFillStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "fillRect",
        argc: 5,
    },
    MethodOp::Return,
];

/// `Graphics.DrawEllipse(pen, x, y, w, h)` — `beginPath` + `ellipse` +
/// `stroke`, WHATWG's own three steps.
///
/// .NET states an ellipse as a BOUNDING BOX and the canvas wants a CENTRE and
/// RADII, so the body converts: `cx = x + w/2`, `cy = y + h/2`, `rx = w/2`,
/// `ry = h/2`.
const GRAPHICS_DRAW_ELLIPSE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    // cx = x + w/2
    MethodOp::PushArg(2),
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // cy = y + h/2
    MethodOp::PushArg(3),
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // rx = w/2
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    // ry = h/2
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "ellipse",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "stroke",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.FillEllipse(brush, x, y, w, h)` — the same conversion as
/// `GRAPHICS_DRAW_ELLIPSE`, filled instead of stroked.
const GRAPHICS_FILL_ELLIPSE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setFillStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    // cx = x + w/2
    MethodOp::PushArg(2),
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // cy = y + h/2
    MethodOp::PushArg(3),
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // rx = w/2
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    // ry = h/2
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "ellipse",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "fill",
        argc: 1,
    },
    MethodOp::Return,
];

// ── The `Rectangle` overloads ────────────────────────────────────────────────
//
// GDI+ states half its drawing surface twice: once as four loose coordinates
// and once as a `Rectangle`. The canvas has ONE shape for each and it takes
// numbers, so these bodies differ from the coordinate forms only in where the
// numbers come from — `rect.width` instead of arg 4. `dotnet.rectangle_new`
// stores lowercase `x`/`y`/`width`/`height`, which is what `PushArgField`
// reads.
//
// Arity 3 (`this` + pen/brush + rect) against the coordinate forms' 6, which is
// what lets `drawing_method_body` tell them apart.

/// `Graphics.DrawRectangle(pen, Rectangle rect)`
const GRAPHICS_DRAW_RECTANGLE_RECT: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(2, "x"),
    MethodOp::PushArgField(2, "y"),
    MethodOp::PushArgField(2, "width"),
    MethodOp::PushArgField(2, "height"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "strokeRect",
        argc: 5,
    },
    MethodOp::Return,
];

/// `Graphics.FillRectangle(brush, Rectangle rect)`
const GRAPHICS_FILL_RECTANGLE_RECT: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setFillStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(2, "x"),
    MethodOp::PushArgField(2, "y"),
    MethodOp::PushArgField(2, "width"),
    MethodOp::PushArgField(2, "height"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "fillRect",
        argc: 5,
    },
    MethodOp::Return,
];

/// `Graphics.DrawEllipse(pen, Rectangle rect)` — the same bounding-box-to-
/// centre-and-radii conversion as the coordinate form, read off the rect.
const GRAPHICS_DRAW_ELLIPSE_RECT: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    // cx = rect.x + rect.width/2
    MethodOp::PushArgField(2, "x"),
    MethodOp::PushArgField(2, "width"),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // cy = rect.y + rect.height/2
    MethodOp::PushArgField(2, "y"),
    MethodOp::PushArgField(2, "height"),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // rx = rect.width/2
    MethodOp::PushArgField(2, "width"),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    // ry = rect.height/2
    MethodOp::PushArgField(2, "height"),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "ellipse",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "stroke",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.FillEllipse(brush, Rectangle rect)`
const GRAPHICS_FILL_ELLIPSE_RECT: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setFillStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(2, "x"),
    MethodOp::PushArgField(2, "width"),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    MethodOp::PushArgField(2, "y"),
    MethodOp::PushArgField(2, "height"),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    MethodOp::PushArgField(2, "width"),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::PushArgField(2, "height"),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "ellipse",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "fill",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.Clear(color)` — fill the whole surface with `color`.
///
/// The canvas has no clear-to-colour operation and does not need one: this IS
/// the idiom, and it is what a browser would run —
/// `setFillStyle` + a `fillRect` over the bitmap's full extent.
///
/// Three details make it correct rather than approximately correct:
/// - `resetTransform`, because `Clear` covers the surface whatever transform
///   the caller left in place. Inside the `save`/`restore` pair, so their
///   transform survives.
/// - `canvasWidth`/`canvasHeight` (`web:html`), which are HTMLCanvasElement's
///   IDL attributes — the BITMAP's size, not the box's. A 640×480 buffer in a
///   320×240 box is the ordinary way to draw at double density, and filling
///   the box would leave three quarters of the surface untouched. `PushThis`
///   serves as the node argument because the context carries `__node`.
/// - The `save`/`restore` pair is balanced, so it leaves the clip baseline
///   (see `GRAPHICS_SET_CLIP`) exactly where it found it.
const GRAPHICS_CLEAR: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "save",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "resetTransform",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "r"),
    MethodOp::PushArgField(1, "g"),
    MethodOp::PushArgField(1, "b"),
    MethodOp::PushArgField(1, "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setFillStyle",
        argc: 5,
    },
    MethodOp::Drop,
    // fillRect(0, 0, canvasWidth(doc, this), canvasHeight(doc, this))
    MethodOp::PushThis,
    MethodOp::PushConstFloat(0.0),
    MethodOp::PushConstFloat(0.0),
    MethodOp::CallHost {
        module: "web:html",
        fn_name: "activeDocument",
        argc: 0,
    },
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:html",
        fn_name: "canvasWidth",
        argc: 2,
    },
    MethodOp::CallHost {
        module: "web:html",
        fn_name: "activeDocument",
        argc: 0,
    },
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:html",
        fn_name: "canvasHeight",
        argc: 2,
    },
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "fillRect",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "restore",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.Dispose()` — no-op (no GDI handle to free).
const GRAPHICS_DISPOSE: &[MethodOp] = &[MethodOp::PushConstNull, MethodOp::Return];

/// `Graphics.DrawArc(pen, x, y, w, h, startAngle, sweepAngle)` —
/// `beginPath` + `arc` + `stroke`.
///
/// Two conversions. The box becomes a centre and radius the same way
/// `GRAPHICS_DRAW_ELLIPSE` converts one, and the ANGLES become radians:
/// GDI+ states degrees, the canvas takes radians (HTML §4.12.5), and both
/// measure clockwise from the positive x-axis, so a scale factor is the whole
/// of the difference. `sweep` is relative where the canvas's second angle is
/// absolute, hence `end = (start + sweep) * k`.
///
/// ⚠ `radius = w/2` ignores `h`: GDI+ sweeps an ELLIPTICAL arc inside the box
/// and the canvas has no elliptical-arc op that reaches the painter —
/// `web:canvas::ellipse` accepts start/end angles and documents that it drops
/// them. A non-square box therefore draws a circular arc of its width — an
/// approximation, stated rather than hidden.
///
/// A negative `sweep` puts `end` before `start`. The canvas resolves that by
/// going clockwise the long way round rather than counter-clockwise, because
/// the `counterclockwise` argument is not passed — the sign would have to be
/// tested, and the DSL has no branch.
const GRAPHICS_DRAW_ARC: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    // cx = x + w/2
    MethodOp::PushArg(2),
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // cy = y + h/2
    MethodOp::PushArg(3),
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // r = w/2
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    // startAngle = start * π/180
    MethodOp::PushArg(6),
    MethodOp::PushConstFloat(DEG_TO_RAD),
    MethodOp::Mul,
    // endAngle = (start + sweep) * π/180
    MethodOp::PushArg(6),
    MethodOp::PushArg(7),
    MethodOp::Add,
    MethodOp::PushConstFloat(DEG_TO_RAD),
    MethodOp::Mul,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "arc",
        argc: 6,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "stroke",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.DrawPie(pen, x, y, w, h, startAngle, sweepAngle)`
///
/// A pie is an arc CLOSED THROUGH ITS CENTRE, which is exactly what the path
/// API says: move to the centre, sweep the arc, close the path. The same box
/// and angle conversions as `GRAPHICS_DRAW_ARC`, and the same ⚠ about a
/// non-square box drawing a circular sweep.
const GRAPHICS_DRAW_PIE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    // moveTo(cx, cy) — the pie's apex.
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    MethodOp::PushArg(3),
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "moveTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    // cx = x + w/2
    MethodOp::PushArg(2),
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // cy = y + h/2
    MethodOp::PushArg(3),
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // r = w/2
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    // startAngle / endAngle, degrees → radians
    MethodOp::PushArg(6),
    MethodOp::PushConstFloat(DEG_TO_RAD),
    MethodOp::Mul,
    MethodOp::PushArg(6),
    MethodOp::PushArg(7),
    MethodOp::Add,
    MethodOp::PushConstFloat(DEG_TO_RAD),
    MethodOp::Mul,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "arc",
        argc: 6,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "closePath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "stroke",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.FillPie(brush, x, y, w, h, startAngle, sweepAngle)` — the same
/// path as `GRAPHICS_DRAW_PIE`, filled instead of stroked.
const GRAPHICS_FILL_PIE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setFillStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    // moveTo(cx, cy) — the pie's apex.
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    MethodOp::PushArg(3),
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "moveTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    // cx = x + w/2
    MethodOp::PushArg(2),
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // cy = y + h/2
    MethodOp::PushArg(3),
    MethodOp::PushArg(5),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    MethodOp::Add,
    // r = w/2
    MethodOp::PushArg(4),
    MethodOp::PushConstFloat(2.0),
    MethodOp::Div,
    // startAngle / endAngle, degrees → radians
    MethodOp::PushArg(6),
    MethodOp::PushConstFloat(DEG_TO_RAD),
    MethodOp::Mul,
    MethodOp::PushArg(6),
    MethodOp::PushArg(7),
    MethodOp::Add,
    MethodOp::PushConstFloat(DEG_TO_RAD),
    MethodOp::Mul,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "arc",
        argc: 6,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "closePath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "fill",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.DrawBezier(pen, x1, y1, x2, y2, x3, y3, x4, y4)` — cubic bezier.
const GRAPHICS_DRAW_BEZIER: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setStrokeStyle",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2), // x1
    MethodOp::PushArg(3), // y1
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "moveTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(4), // cx1
    MethodOp::PushArg(5), // cy1
    MethodOp::PushArg(6), // cx2
    MethodOp::PushArg(7), // cy2
    MethodOp::PushArg(8), // x4
    MethodOp::PushArg(9), // y4
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "bezierCurveTo",
        argc: 7,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "stroke",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.DrawString(text, font, brush, x, y)`
const GRAPHICS_DRAW_STRING: &[MethodOp] = &[
    // Fill colour from brush.color (arg 3).
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(3, "color", "r"),
    MethodOp::PushArgFieldField(3, "color", "g"),
    MethodOp::PushArgFieldField(3, "color", "b"),
    MethodOp::PushArgFieldField(3, "color", "a"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setFillStyle",
        argc: 5,
    },
    MethodOp::Drop,
    // Font from arg 2 (Font object with name/size/bold/italic).
    MethodOp::PushThis,
    MethodOp::PushArgField(2, "name"),
    MethodOp::PushArgField(2, "size"),
    MethodOp::PushArgField(2, "bold"),
    MethodOp::PushArgField(2, "italic"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setFont",
        argc: 5,
    },
    MethodOp::Drop,
    // `fillText`'s `y` is the BASELINE — `textBaseline` defaults to
    // `alphabetic` (HTML §4.12.5) — while GDI+ `DrawString(s, font, brush, x,
    // y)` positions the text's TOP-LEFT. Passing `y` straight through drew
    // every string about one ascent too high.
    //
    // `textBaseline = "top"` says exactly that in the canvas's own vocabulary,
    // so `y` then means what .NET means by it. The alternative — adding an
    // ascent to `y` — would need `measureText`, which `web:canvas` does not
    // have, and would only ever approximate what this states outright.
    MethodOp::PushThis,
    MethodOp::PushConstStr("top"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "setTextBaseline",
        argc: 2,
    },
    MethodOp::Drop,
    // FillText(text, x, y).
    MethodOp::PushThis,
    MethodOp::PushArg(1), // text
    MethodOp::PushArg(4), // x
    MethodOp::PushArg(5), // y
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "fillText",
        argc: 4,
    },
    MethodOp::Return,
];

/// `Graphics.Save()` — push state. .NET returns a `GraphicsState` token; we
/// return null (the canvas save/restore stack is implicit).
const GRAPHICS_SAVE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "save",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.Restore(state)` — pop state. The `state` arg is ignored (the
/// canvas has a single implicit stack).
const GRAPHICS_RESTORE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "restore",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.TranslateTransform(dx, dy)`
const GRAPHICS_TRANSLATE_TRANSFORM: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArg(1),
    MethodOp::PushArg(2),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "translate",
        argc: 3,
    },
    MethodOp::Return,
];

/// `Graphics.RotateTransform(angleDegrees)` — `rotate(angle)`, HTML §4.12.5.
///
/// The canvas takes RADIANS, .NET states degrees, so the caller converts. A
/// host function whose whole job is one multiplication is a host function a
/// browser engine could not be asked for.
const GRAPHICS_ROTATE_TRANSFORM: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArg(1),
    MethodOp::PushConstFloat(DEG_TO_RAD),
    MethodOp::Mul,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "rotate",
        argc: 2,
    },
    MethodOp::Return,
];

/// `Graphics.ScaleTransform(sx, sy)`
const GRAPHICS_SCALE_TRANSFORM: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArg(1),
    MethodOp::PushArg(2),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "scale",
        argc: 3,
    },
    MethodOp::Return,
];

/// `Graphics.ResetTransform()`
const GRAPHICS_RESET_TRANSFORM: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "resetTransform",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.SetClip(x, y, w, h)` — rect form (the most common overload).
///
/// **The clip baseline.** A canvas clip only ever INTERSECTS — `clip()` has no
/// inverse — while .NET's `SetClip` REPLACES the region and `ResetClip` removes
/// it. `save`/`restore` are what bridge the two, because they push and pop the
/// clip along with the rest of the paint state (`widgets::canvas::Canvas`
/// says so on `clip()`), so a region is undone by returning to a state saved
/// before it was applied.
///
/// So a `Graphics` holds exactly one baseline entry on the canvas state stack
/// for its whole life: `CreateGraphics` pushes it right after `getContext`, and
/// every clip operation pops back to it and re-pushes it. Depth is invariant —
/// `restore` here can never underflow, because the baseline is always there.
///
/// ⚠ `Graphics.Save()`/`Restore()` push and pop that same stack, so a `SetClip`
/// between a Save and its Restore pops the Save's entry instead of the clip
/// baseline. Matching .NET exactly there needs a clip region per saved state,
/// which the canvas cannot express and a stateless `Body` cannot track. Flat
/// use — the shape essentially every WinForms program has — is exact.
const GRAPHICS_SET_CLIP: &[MethodOp] = &[
    // Back to the baseline: drops whatever region a previous SetClip applied.
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "restore",
        argc: 1,
    },
    MethodOp::Drop,
    // Re-establish it, so the next SetClip/ResetClip has one to return to.
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "save",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "beginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(1),
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "rect",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "clip",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.ResetClip()` — return to the clip baseline and re-establish it.
///
/// WHATWG has no reset-clip operation and does not need one: a clip is undone
/// by returning to a state saved before it was applied. See
/// `GRAPHICS_SET_CLIP` for the baseline this pops back to and the one case it
/// does not model.
const GRAPHICS_RESET_CLIP: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "restore",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "save",
        argc: 1,
    },
    MethodOp::Return,
];

const GRAPHICS_METHODS: &[DotnetMethod] = &[
    DotnetMethod {
        name: "DrawLine",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_DRAW_LINE),
    },
    // The two-`Point` overload — `this` + pen + 2 points = 4. Listed beside the
    // coordinate form rather than replacing it: `drawing_method_body` picks by
    // arity, so both spellings resolve.
    DotnetMethod {
        name: "DrawLine",
        arity: 4,
        target: MethodTarget::body(GRAPHICS_DRAW_LINE_POINTS),
    },
    DotnetMethod {
        name: "DrawRectangle",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_DRAW_RECTANGLE),
    },
    DotnetMethod {
        name: "DrawRectangle",
        arity: 3,
        target: MethodTarget::body(GRAPHICS_DRAW_RECTANGLE_RECT),
    },
    DotnetMethod {
        name: "DrawEllipse",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_DRAW_ELLIPSE),
    },
    DotnetMethod {
        name: "DrawEllipse",
        arity: 3,
        target: MethodTarget::body(GRAPHICS_DRAW_ELLIPSE_RECT),
    },
    DotnetMethod {
        name: "DrawArc",
        arity: 8,
        target: MethodTarget::body(GRAPHICS_DRAW_ARC),
    },
    DotnetMethod {
        name: "DrawPie",
        arity: 8,
        target: MethodTarget::body(GRAPHICS_DRAW_PIE),
    },
    DotnetMethod {
        name: "FillPie",
        arity: 8,
        target: MethodTarget::body(GRAPHICS_FILL_PIE),
    },
    DotnetMethod {
        name: "DrawBezier",
        arity: 10,
        target: MethodTarget::body(GRAPHICS_DRAW_BEZIER),
    },
    DotnetMethod {
        name: "DrawString",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_DRAW_STRING),
    },
    DotnetMethod {
        name: "FillRectangle",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_FILL_RECTANGLE),
    },
    DotnetMethod {
        name: "FillRectangle",
        arity: 3,
        target: MethodTarget::body(GRAPHICS_FILL_RECTANGLE_RECT),
    },
    DotnetMethod {
        name: "FillEllipse",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_FILL_ELLIPSE),
    },
    DotnetMethod {
        name: "FillEllipse",
        arity: 3,
        target: MethodTarget::body(GRAPHICS_FILL_ELLIPSE_RECT),
    },
    DotnetMethod {
        name: "Clear",
        arity: 2,
        target: MethodTarget::body(GRAPHICS_CLEAR),
    },
    DotnetMethod {
        name: "Save",
        arity: 1,
        target: MethodTarget::body(GRAPHICS_SAVE),
    },
    DotnetMethod {
        name: "Restore",
        arity: 2,
        target: MethodTarget::body(GRAPHICS_RESTORE),
    },
    DotnetMethod {
        name: "TranslateTransform",
        arity: 3,
        target: MethodTarget::body(GRAPHICS_TRANSLATE_TRANSFORM),
    },
    DotnetMethod {
        name: "RotateTransform",
        arity: 2,
        target: MethodTarget::body(GRAPHICS_ROTATE_TRANSFORM),
    },
    DotnetMethod {
        name: "ScaleTransform",
        arity: 3,
        target: MethodTarget::body(GRAPHICS_SCALE_TRANSFORM),
    },
    DotnetMethod {
        name: "ResetTransform",
        arity: 1,
        target: MethodTarget::body(GRAPHICS_RESET_TRANSFORM),
    },
    DotnetMethod {
        name: "SetClip",
        arity: 5,
        target: MethodTarget::body(GRAPHICS_SET_CLIP),
    },
    DotnetMethod {
        name: "ResetClip",
        arity: 1,
        target: MethodTarget::body(GRAPHICS_RESET_CLIP),
    },
    DotnetMethod {
        name: "Dispose",
        arity: 1,
        target: MethodTarget::body(GRAPHICS_DISPOSE),
    },
];

/// `Pen.Dispose()` and `Brush.Dispose()` — no-op for now.
const PEN_METHODS: &[DotnetMethod] = &[DotnetMethod {
    name: "Dispose",
    arity: 1,
    target: MethodTarget::body(GRAPHICS_DISPOSE),
}];
const BRUSH_METHODS: &[DotnetMethod] = &[DotnetMethod {
    name: "Dispose",
    arity: 1,
    target: MethodTarget::body(GRAPHICS_DISPOSE),
}];

/// A named `Color` static — `Color.Red` IS its four channels.
///
/// The RGBA is pushed and the object composed by `MethodOp::NewValueType`, the
/// same `emit_value_type_new` `Point` and `Size` reach through their
/// constructors — NOT `Color`'s own ctor, which would be circular for a static
/// declared on `Color`. It used to push the colour NAME and look it up in a
/// host-side palette — a round-trip to turn a compile-time constant into four
/// compile-time constants. Nothing about a colour needs a
/// host: it is data, and this is where the data belongs now that it is stated
/// once.
///
/// Still a body rather than a `Fn` leaf, for the original reason:
/// `ResolutionTarget::HostCall` carries no bound argument — `terminal()`
/// discards `NamespaceNode::Fn`'s `bound_arg` — so a leaf would answer with no
/// arguments at all.
/// A `Color`'s fields, in the order every builder pushes them.
///
/// Single-sourced because two places compose a `Color` — `dotnet.color_new`
/// (`New Color(r, g, b, a)`) and the named statics below — and the drawing
/// bodies read the result BY NAME (`PushArgFieldField(1, "color", "r")`). If
/// the two lists drifted, `Color.Red` and `New Color(220, 20, 60, 255)` would
/// be different objects that both look right.
pub(crate) const COLOR_FIELDS: &[&str] = &["r", "g", "b", "a"];

macro_rules! color_static {
    ($konst:ident, $r:literal, $g:literal, $b:literal, $a:literal) => {
        const $konst: &[MethodOp] = &[
            // **A COLOUR IS DATA.** Four numbers in an object, composed in
            // bytecode — there is no host in it and no web API for it either,
            // because a colour is not a browser capability.
            //
            // This was the LAST host import any .NET program still carried. It
            // stayed a host call through two failed attempts, both
            // recorded here because the reason is not obvious:
            //   - `NewDotnet { class: "Color", argc: 4 }` — `builder.rs`
            //     asserts `argc == 0`; the DSL has no arity-N factory.
            //   - `NewDotnet { argc: 0 }` + `Dup`/`SetField` per channel —
            //     CIRCULAR. `Color.Red` is a static ON `Color`, and
            //     `NewDotnet` reads `Color`'s ctor GLOBAL, installed by an
            //     EARLIER registration pass, so it answers `undefined`.
            // `NewValueType` composes the object with no constructor at all,
            // which is what removes the ordering dependency.
            //
            // ⚠ Field ORDER is the contract: `emit_value_type_new` pops one
            // value per field, last-first, so these push in `COLOR_FIELDS`
            // order — `(r, g, b, a)`. The retired host fn was
            // `color.fromargb(a, r, g, b)`, a DIFFERENT order, which is why
            // this is a reordering and not just a swapped call.
            MethodOp::PushConstFloat($r as f64),
            MethodOp::PushConstFloat($g as f64),
            MethodOp::PushConstFloat($b as f64),
            MethodOp::PushConstFloat($a as f64),
            MethodOp::NewValueType {
                type_name: "Color",
                fields: COLOR_FIELDS,
            },
            MethodOp::Return,
        ];
    };
}

color_static!(COLOR_RED, 220, 20, 60, 255);
color_static!(COLOR_BLUE, 30, 144, 255, 255);
color_static!(COLOR_GREEN, 34, 139, 34, 255);
color_static!(COLOR_BLACK, 0, 0, 0, 255);
color_static!(COLOR_WHITE, 255, 255, 255, 255);
color_static!(COLOR_YELLOW, 255, 215, 0, 255);
color_static!(COLOR_ORANGE, 255, 140, 0, 255);
color_static!(COLOR_PURPLE, 128, 0, 128, 255);
color_static!(COLOR_CYAN, 0, 255, 255, 255);
color_static!(COLOR_MAGENTA, 255, 0, 255, 255);
color_static!(COLOR_GRAY, 128, 128, 128, 255);
color_static!(COLOR_BROWN, 139, 69, 19, 255);
color_static!(COLOR_PINK, 255, 192, 203, 255);
color_static!(COLOR_LIGHT_GRAY, 211, 211, 211, 255);
color_static!(COLOR_DARK_GRAY, 169, 169, 169, 255);
color_static!(COLOR_TRANSPARENT, 0, 0, 0, 0);

/// The `Color` statics, as (member name, body). `tree_register` registers each
/// one at `dotnet.system.drawing.color.<lowercased>` — a real path segment, so
/// `merge_into`'s `Type{statics} × Namespace` arm folds it into the already
/// registered `Color` type's statics. A static IS a longer path.
pub const COLOR_STATICS: &[(&str, &[MethodOp])] = &[
    ("Red", COLOR_RED),
    ("Blue", COLOR_BLUE),
    ("Green", COLOR_GREEN),
    ("Black", COLOR_BLACK),
    ("White", COLOR_WHITE),
    ("Yellow", COLOR_YELLOW),
    ("Orange", COLOR_ORANGE),
    ("Purple", COLOR_PURPLE),
    ("Cyan", COLOR_CYAN),
    ("Magenta", COLOR_MAGENTA),
    ("Gray", COLOR_GRAY),
    ("Brown", COLOR_BROWN),
    ("Pink", COLOR_PINK),
    ("LightGray", COLOR_LIGHT_GRAY),
    ("DarkGray", COLOR_DARK_GRAY),
    ("Transparent", COLOR_TRANSPARENT),
];

/// The `dotnet.drawing.*` emit key for a `Color` static, as registered in the
/// tree and as matched by [`drawing_method_body`]. One function so the
/// registrar and the dispatcher cannot disagree about the spelling.
pub fn color_static_emit_key(member: &str) -> String {
    format!("dotnet.drawing.color.{}", member.to_lowercase())
}

/// Look up a drawing method's `Body` ops by name AND arity across the
/// Graphics/Pen/Brush tables. The `dotnet.drawing.*` call-site dispatch uses
/// this to lower the method inline (`builder::emit_body_inline`) — the drawing
/// objects resolve their methods through the component descriptor
/// (`MethodBody::Common`) with no ctor-bound thunk, the same way controls
/// resolve theirs.
///
/// **`argc` is what makes overloads expressible, and it is not optional.** The
/// dispatch key is `dotnet.drawing.<method>` — a name with no arity in it — so
/// keying the tables by name alone gave a method ONE body, whichever entry came
/// first. GDI+ overloads heavily: `DrawLine` is both `(pen, x1, y1, x2, y2)` and
/// `(pen, Point, Point)`, `DrawEllipse` is both `(pen, x, y, w, h)` and
/// `(pen, Rectangle)`. A narrower call then reached the wider body and its
/// `PushArg(n)` ran off the end of the frame — `builder.rs` PANICS on that, so
/// `g.DrawLine(Pens.Black, p1, p2)` killed the whole compile rather than missing
/// (`examples/vb/paint_demo`).
///
/// `arity` counts `this`, so the two `DrawLine`s are 6 and 4. An exact match
/// wins; failing that a narrower body is accepted for a wider call, which is
/// how a method with trailing optional arguments still resolves. A body WIDER
/// than the call is never returned — that is the panic, and answering `None`
/// makes it an ordinary unresolved method instead.
pub fn drawing_method_body(name: &str, argc: u8) -> Option<&'static [MethodOp]> {
    // `Color` statics are keyed `color.<lowercased>` — a member READ, argc 0,
    // not an instance method, so it cannot collide with the method tables
    // below (none of which contain a dot).
    if let Some(member) = name.strip_prefix("color.") {
        for (candidate, ops) in COLOR_STATICS {
            if candidate.to_lowercase() == member {
                return Some(ops);
            }
        }
        return None;
    }
    let mut widest_narrower: Option<(u8, &'static [MethodOp])> = None;
    for table in [GRAPHICS_METHODS, PEN_METHODS, BRUSH_METHODS] {
        for m in table {
            if m.name != name {
                continue;
            }
            let MethodTarget::Body(ops) = m.target else {
                continue;
            };
            if m.arity == argc {
                return Some(ops);
            }
            // Narrower than the call: reachable, because every `PushArg` it
            // names is inside the frame. Keep the widest such body so a
            // trailing-optional call still lands on the closest overload.
            if m.arity < argc && widest_narrower.is_none_or(|(best, _)| m.arity > best) {
                widest_narrower = Some((m.arity, ops));
            }
        }
    }
    widest_narrower.map(|(_, ops)| ops)
}

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "Graphics",
            parent: Some("MarshalByRefObject"),
            properties: &[
                "Clip",
                "ClipBounds",
                "CompositingMode",
                "CompositingQuality",
                "DpiX",
                "DpiY",
                "InterpolationMode",
                "IsClipEmpty",
                "PageScale",
                "PageUnit",
                "PixelOffsetMode",
                "RenderingOrigin",
                "SmoothingMode",
                "TextContrast",
                "TextRenderingHint",
                "Transform",
                "VisibleClipBounds",
            ],
            methods: GRAPHICS_METHODS,
            ctor_arity: 0,
            // A `Graphics` is normally created by `Control.CreateGraphics()`,
            // which is `element.getContext("2d")` and bypasses this ctor
            // entirely — the surface belongs to the control it came from.
            //
            // ⚠ This `widget_host_fn` backs bare `New Graphics()`, which has no
            // element and therefore no surface: its drawing calls resolve
            // nothing and paint nothing. .NET has no such constructor either
            // (`Graphics` is obtained `FromImage`/`FromHwnd`/`CreateGraphics`),
            // so the honest answer is a `Graphics` over an offscreen canvas —
            // which needs an element this descriptor has no way to make.
            // Composed by `dotnet.graphics_new` — an identity record. A real
            // surface comes from `CreateGraphics`/`FromImage`, never from here.
            widget_host_fn: None,        },
        DotnetClass {
            name: "Pen",
            parent: Some("MarshalByRefObject"),
            properties: &[
                "Alignment",
                "Brush",
                "Color",
                "CompoundArray",
                "CustomEndCap",
                "CustomStartCap",
                "DashCap",
                "DashOffset",
                "DashPattern",
                "DashStyle",
                "EndCap",
                "LineJoin",
                "MiterLimit",
                "PenType",
                "StartCap",
                "Transform",
                "Width",
            ],
            methods: PEN_METHODS,
            ctor_arity: 2,
            // Composed by `dotnet.pen_new` — a record of colour and width.
            widget_host_fn: None,        },
        DotnetClass {
            name: "Brush",
            parent: Some("MarshalByRefObject"),
            properties: &[],
            methods: BRUSH_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "SolidBrush",
            parent: Some("Brush"),
            properties: &["Color"],
            methods: &[],
            ctor_arity: 1,
            // Composed by `dotnet.solid_brush_new` — a record of one colour.
            widget_host_fn: None,        },
        DotnetClass {
            name: "HatchBrush",
            parent: Some("Brush"),
            properties: &["BackgroundColor", "ForegroundColor", "HatchStyle"],
            methods: &[],
            ctor_arity: 3,
            // Composed by `dotnet.hatch_brush_new`.
            widget_host_fn: None,        },
        DotnetClass {
            name: "LinearGradientBrush",
            parent: Some("Brush"),
            properties: &[
                "Blend",
                "GammaCorrection",
                "InterpolationColors",
                "LinearColors",
                "Rectangle",
                "Transform",
                "WrapMode",
            ],
            methods: &[],
            ctor_arity: 4,
            // Composed by `dotnet.linear_gradient_brush_new`.
            widget_host_fn: None,        },
        // System.Drawing.Point — position value type. `new Point(x, y)`
        // composes an object with `{x, y}` in bytecode (`dotnet.point_new`).
        // `.X` / `.Y` are NOT accessors: a value type declares none, so they
        // resolve as ordinary struct field reads on the lowercased names.
        DotnetClass {
            name: "Point",
            parent: None,
            properties: &["X", "Y", "IsEmpty"],
            methods: &[],
            ctor_arity: 2,
            // Composed by `dotnet.point_new` — see `common_ctor_for`. A value
            // type has no element, so the factory only ever allocated.
            widget_host_fn: None,        },
        // System.Drawing.Size — dimensions value type. Mirror of Point.
        DotnetClass {
            name: "Size",
            parent: None,
            properties: &["Width", "Height", "IsEmpty"],
            methods: &[],
            ctor_arity: 2,
            // Composed by `dotnet.size_new` — the mirror of `Point`.
            widget_host_fn: None,        },
        // System.Drawing.Color — four channels, 0-255. The `Graphics` bodies
        // read `pen.color.r`/`.g`/`.b`/`.a` as NUMBERS to build
        // `web:canvas::setStrokeStyle`, so the channel names are the contract.
        DotnetClass {
            name: "Color",
            parent: None,
            properties: &["R", "G", "B", "A", "Name", "IsEmpty"],
            methods: &[],
            ctor_arity: 4,
            // Composed by `dotnet.color_new` — see `common_ctor_for`.
            widget_host_fn: None,        },
        // System.Drawing.Font — `New Font(name, size)`.
        //
        // It had NO class entry here at all: its only declaration was a host
        // factory backing, so it was the last type still built that way.
        // `Bold`/`Italic`
        // are listed because the drawing bodies read them off a font argument;
        // the two-argument overload leaves both false, which is the contract
        // the retired factory set.
        DotnetClass {
            name: "Font",
            parent: None,
            properties: &["Name", "Size", "Bold", "Italic"],
            methods: &[],
            ctor_arity: 2,
            // Composed by `dotnet.font_new` — see `common_ctor_for`.
            widget_host_fn: None,        },
        // System.Drawing.Rectangle — position AND extent in one value, and the
        // argument GDI+ overloads half its drawing surface on
        // (`DrawEllipse(pen, rect)`, `DrawRectangle(pen, rect)`).
        //
        // It was never declared, so `New Rectangle(x, y, w, h)` answered
        // `undefined is not callable` and `examples/vb/paint_demo` could not
        // draw a circle or a rectangle at all.
        //
        // ⚠ NO `widget_host_fn` — a rectangle is four numbers in an object and
        // needs nothing from a host.
        // Its constructor is `dotnet.rectangle_new`, composed from primitives
        // (`common_ctor_for` in `winforms/component_classes.rs`), which is the
        // route `pointNew`/`sizeNew` should follow next.
        //
        // `Left`/`Top`/`Right`/`Bottom` are DERIVED in .NET (`Right = X +
        // Width`) and are not stored; they are listed so the property axis
        // knows the names, and read back through the same keyed getter that
        // answers `Point.X`.
            // **The floating-point mirrors — `SizeF`, `PointF`, `RectangleF`.**
            //
            // `System.Drawing` declares each value type twice, once in integers
            // and once in `Single`, and only the integer half was here. Every
            // VB designer emits
            // `Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)`,
            // so a form whose designer includes that line failed with
            // "undefined is not callable" inside `InitializeComponent` — the
            // constructor resolved to nothing. Samples that happened to omit
            // the line ran, which is why this looked like one broken project
            // rather than a missing type.
            //
            // Separate declarations rather than aliases of the integer ones:
            // the `__type` string IS the identity a value type compares by, so
            // a `SizeF` answering "Size" would make the two equal.
            DotnetClass {
                name: "SizeF",
                parent: None,
                properties: &["Width", "Height", "IsEmpty"],
                methods: &[],
                ctor_arity: 2,
                widget_host_fn: None,
            },
            DotnetClass {
                name: "PointF",
                parent: None,
                properties: &["X", "Y", "IsEmpty"],
                methods: &[],
                ctor_arity: 2,
                widget_host_fn: None,
            },
            DotnetClass {
                name: "RectangleF",
                parent: None,
                properties: &[
                    "X", "Y", "Width", "Height", "Left", "Top", "Right", "Bottom", "Location",
                    "Size", "IsEmpty",
                ],
                methods: &[],
                ctor_arity: 4,
                widget_host_fn: None,
            },
        DotnetClass {
            name: "Rectangle",
            parent: None,
            properties: &[
                "X", "Y", "Width", "Height", "Left", "Top", "Right", "Bottom", "Location", "Size",
                "IsEmpty",
            ],
            methods: &[],
            ctor_arity: 4,
            widget_host_fn: None,        },
    ]
}
