//! `System.Drawing` value-shaped types: `Graphics`, `Pen`, `Brush`,
//! `SolidBrush`, `HatchBrush`, `LinearGradientBrush`.
//!
//! ## Architecture
//!
//! `Graphics` looks and acts exactly like .NET's `System.Drawing.Graphics`,
//! but every drawing call routes through the generic `vybe:gui::canvas*`
//! host bridge — which talks to the `vybe_widgets::canvas::Canvas` trait.
//! That trait is HTML5-canvas-shaped, so the .NET method bodies translate
//! the .NET API ("DrawLine takes a Pen with a Color + Width") into the
//! canvas API ("set stroke style, set line width, begin path, move to,
//! line to, stroke").
//!
//! Each `Graphics` method is a [`MethodTarget::Body`] sequence — a small
//! declarative slice of [`MethodOp`]s the builder compiles to bytecode.
//! The body reads `pen.color.r/g/b/a` and `pen.width` from the user's
//! arguments via `PushArgField` and forwards them to the canvas host fns.
//!
//! `Control.CreateGraphics()` is also a `Body` — it calls
//! `vybe:gui::createGraphics(this.__control_name)` which returns a
//! Graphics-shaped Object stamped with the source control's name. The
//! drawing host fns then use that name to find the target
//! `RecordingCanvas` (either a Canvas widget's own recording or an
//! overlay recording on `GuiState`).
//!
//! ## Pen / Brush
//!
//! `Pen` and `SolidBrush` are real dotnet classes with arity-N
//! constructors. The user writes `New Pen(Color.Red, 5)` and the dotnet
//! ctor forwards to `vybe:gui::penNew(color, width)` — which returns
//! an Object with `color` and `width` fields. The Graphics method bodies
//! read those fields directly.
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
        module: "vybe:gui",
        fn_name: "canvasSetStrokeColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasBeginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasMoveTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasLineTo",
        argc: 3,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasStroke",
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
        module: "vybe:gui",
        fn_name: "canvasSetStrokeColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasStrokeRect",
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
        module: "vybe:gui",
        fn_name: "canvasSetFillColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasFillRect",
        argc: 5,
    },
    MethodOp::Return,
];

/// `Graphics.DrawEllipse(pen, x, y, w, h)`
///
/// Translates the .NET "bounding box" form to the canvas "centre + radii"
/// form: `cx = x + w/2`, `cy = y + h/2`, `rx = w/2`, `ry = h/2`. We
/// can't do arithmetic in the body DSL (yet), so we instead use the
/// canvas's `rect` op to build a stroked rectangle that approximates the
/// ellipse's bounding box. A future extension can add an arithmetic op
/// or expose `canvasEllipseFromBounds(x, y, w, h)` as a host shortcut.
///
/// For now: stroke the bounding rectangle. Visually approximate but
/// matches the API shape.
///
/// TODO: add `canvasEllipseFromBounds` host fn or arithmetic ops.
const GRAPHICS_DRAW_ELLIPSE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetStrokeColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    // Use `canvasEllipseFromBounds` — a host helper that does the
    // x+w/2, y+h/2, w/2, h/2 conversion + begin_path + ellipse +
    // stroke in one call. See `vybe_host::modules::canvas`.
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasStrokeEllipseInRect",
        argc: 5,
    },
    MethodOp::Return,
];

/// `Graphics.FillEllipse(brush, x, y, w, h)`
const GRAPHICS_FILL_ELLIPSE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetFillColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasFillEllipseInRect",
        argc: 5,
    },
    MethodOp::Return,
];

/// `Graphics.Clear(color)` — clears the entire canvas to a colour.
const GRAPHICS_CLEAR: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "r"),
    MethodOp::PushArgField(1, "g"),
    MethodOp::PushArgField(1, "b"),
    MethodOp::PushArgField(1, "a"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasClearAll",
        argc: 5,
    },
    MethodOp::Return,
];

/// `Graphics.Dispose()` — no-op (no GDI handle to free).
const GRAPHICS_DISPOSE: &[MethodOp] = &[MethodOp::PushConstNull, MethodOp::Return];

/// `Graphics.DrawArc(pen, x, y, w, h, startAngle, sweepAngle)`
const GRAPHICS_DRAW_ARC: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetStrokeColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2), // x
    MethodOp::PushArg(3), // y
    MethodOp::PushArg(4), // w
    MethodOp::PushArg(5), // h
    MethodOp::PushArg(6), // startAngle (deg)
    MethodOp::PushArg(7), // sweepAngle (deg)
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasStrokeArcInRect",
        argc: 7,
    },
    MethodOp::Return,
];

/// `Graphics.DrawPie(pen, x, y, w, h, startAngle, sweepAngle)`
const GRAPHICS_DRAW_PIE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetStrokeColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::PushArg(6),
    MethodOp::PushArg(7),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasStrokePieInRect",
        argc: 7,
    },
    MethodOp::Return,
];

/// `Graphics.FillPie(brush, x, y, w, h, startAngle, sweepAngle)`
const GRAPHICS_FILL_PIE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArgFieldField(1, "color", "r"),
    MethodOp::PushArgFieldField(1, "color", "g"),
    MethodOp::PushArgFieldField(1, "color", "b"),
    MethodOp::PushArgFieldField(1, "color", "a"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetFillColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::PushArg(5),
    MethodOp::PushArg(6),
    MethodOp::PushArg(7),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasFillPieInRect",
        argc: 7,
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
        module: "vybe:gui",
        fn_name: "canvasSetStrokeColor",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArgField(1, "width"),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSetLineWidth",
        argc: 2,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasBeginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(2), // x1
    MethodOp::PushArg(3), // y1
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasMoveTo",
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
        module: "vybe:gui",
        fn_name: "canvasBezierTo",
        argc: 7,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasStroke",
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
        module: "vybe:gui",
        fn_name: "canvasSetFillColor",
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
        module: "vybe:gui",
        fn_name: "canvasSetFont",
        argc: 5,
    },
    MethodOp::Drop,
    // FillText(text, x, y).
    MethodOp::PushThis,
    MethodOp::PushArg(1), // text
    MethodOp::PushArg(4), // x
    MethodOp::PushArg(5), // y
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasFillText",
        argc: 4,
    },
    MethodOp::Return,
];

/// `Graphics.Save()` — push state. .NET returns a `GraphicsState` token; we
/// return null (the canvas save/restore stack is implicit).
const GRAPHICS_SAVE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasSave",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.Restore(state)` — pop state. The `state` arg is ignored (the
/// canvas has a single implicit stack).
const GRAPHICS_RESTORE: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasRestore",
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
        module: "vybe:gui",
        fn_name: "canvasTranslate",
        argc: 3,
    },
    MethodOp::Return,
];

/// `Graphics.RotateTransform(angleDegrees)`
const GRAPHICS_ROTATE_TRANSFORM: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushArg(1),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasRotateDegrees",
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
        module: "vybe:gui",
        fn_name: "canvasScale",
        argc: 3,
    },
    MethodOp::Return,
];

/// `Graphics.ResetTransform()`
const GRAPHICS_RESET_TRANSFORM: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasResetTransform",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.SetClip(x, y, w, h)` — rect form (the most common overload).
const GRAPHICS_SET_CLIP: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasBeginPath",
        argc: 1,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::PushArg(1),
    MethodOp::PushArg(2),
    MethodOp::PushArg(3),
    MethodOp::PushArg(4),
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasRect",
        argc: 5,
    },
    MethodOp::Drop,
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasClip",
        argc: 1,
    },
    MethodOp::Return,
];

/// `Graphics.ResetClip()`
const GRAPHICS_RESET_CLIP: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::CallHost {
        module: "vybe:gui",
        fn_name: "canvasResetClip",
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
    DotnetMethod {
        name: "DrawRectangle",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_DRAW_RECTANGLE),
    },
    DotnetMethod {
        name: "DrawEllipse",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_DRAW_ELLIPSE),
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
        name: "FillEllipse",
        arity: 6,
        target: MethodTarget::body(GRAPHICS_FILL_ELLIPSE),
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

/// Look up a drawing method's `Body` ops by name across the Graphics/Pen/Brush
/// tables. The `dotnet.drawing.*` call-site dispatch uses this to lower the
/// method inline (`builder::emit_body_inline`) — the drawing objects resolve
/// their methods through the component descriptor (`MethodBody::Common`) with
/// no ctor-bound thunk, the same way controls resolve theirs.
pub fn drawing_method_body(name: &str) -> Option<&'static [MethodOp]> {
    for table in [GRAPHICS_METHODS, PEN_METHODS, BRUSH_METHODS] {
        for m in table {
            if m.name == name {
                if let MethodTarget::Body(ops) = m.target {
                    return Some(ops);
                }
            }
        }
    }
    None
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
            // `Graphics` instances are normally created via
            // `Control.CreateGraphics()` which is a Body that calls
            // `vybe:gui::createGraphics(name)` directly — bypassing this
            // ctor. The `widget_host_fn` here is a fallback for direct
            // `New Graphics()` use, which produces a Graphics object
            // stamped with `__control_name = "graphics"` (a default
            // global canvas, useful for ad-hoc drawing in tests).
            widget_host_fn: Some("graphicsNew"),
            widget_host_module: "vybe:gui",
        },
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
            widget_host_fn: Some("penNew"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "Brush",
            parent: Some("MarshalByRefObject"),
            properties: &[],
            methods: BRUSH_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "SolidBrush",
            parent: Some("Brush"),
            properties: &["Color"],
            methods: &[],
            ctor_arity: 1,
            widget_host_fn: Some("solidBrushNew"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "HatchBrush",
            parent: Some("Brush"),
            properties: &["BackgroundColor", "ForegroundColor", "HatchStyle"],
            methods: &[],
            ctor_arity: 3,
            widget_host_fn: Some("hatchBrushNew"),
            widget_host_module: "vybe:gui",
        },
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
            widget_host_fn: Some("linearGradientBrushNew"),
            widget_host_module: "vybe:gui",
        },
        // System.Drawing.Point — position value type. `new Point(x, y)`
        // lowers to `vybe:gui::pointNew(x, y)` which returns an
        // Object with `{x, y, X, Y}` fields. The GUI property dispatch
        // reads `.x` / `.y` from a control's `location` property.
        DotnetClass {
            name: "Point",
            parent: None,
            properties: &["X", "Y", "IsEmpty"],
            methods: &[],
            ctor_arity: 2,
            widget_host_fn: Some("pointNew"),
            widget_host_module: "vybe:gui",
        },
        // System.Drawing.Size — dimensions value type. Mirror of Point.
        DotnetClass {
            name: "Size",
            parent: None,
            properties: &["Width", "Height", "IsEmpty"],
            methods: &[],
            ctor_arity: 2,
            widget_host_fn: Some("sizeNew"),
            widget_host_module: "vybe:gui",
        },
    ]
}
