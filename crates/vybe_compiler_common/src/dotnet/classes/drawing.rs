//! `System.Drawing` value-shaped types: `Graphics`, `Pen`, `Brush`,
//! `SolidBrush`, `HatchBrush`, `LinearGradientBrush`.
//!
//! Unlike WinForms controls these don't inherit from `Control` — they
//! inherit straight from `Object` (technically `Brush` and `Pen` inherit
//! from `MarshalByRefObject` in real .NET, but the distinction doesn't
//! matter for our class registry). They have their own host backing in
//! `vybe:drawing` and a small set of methods that user code calls
//! directly:
//!
//! - `Graphics.DrawLine(pen, x1, y1, x2, y2)` etc. — drawing primitives
//! - `Pen.Dispose()`, `Brush.Dispose()` — VB6/VB.NET resource cleanup
//!
//! `Pen` and `SolidBrush` have arity-N constructors so user code can
//! write `New Pen(Color.Red, 5)` and `New SolidBrush(Color.Blue)`. The
//! ctor args are forwarded straight to the backing host fn (`penNew` /
//! `solidBrushNew`).
//!
//! Color, Point, Size, and Font remain registered via the legacy
//! `vybe_host::namespaces::*` constructors for now — they're "value
//! types" with no methods worth wrapping in a class.

use super::{DotnetClass, DotnetMethod, MethodTarget};

/// Methods bound on every `Graphics` instance. Each method takes `this`
/// as its first arg (the host fn signature uses `args[0]` to find the
/// graphics target) and forwards the user-supplied args after it.
const GRAPHICS_METHODS: &[DotnetMethod] = &[
    // Pen-based drawing primitives.
    DotnetMethod { name: "DrawLine",      arity: 6, target: MethodTarget::host("vybe:drawing", "drawLine") },
    DotnetMethod { name: "DrawRectangle", arity: 6, target: MethodTarget::host("vybe:drawing", "drawRectangle") },
    DotnetMethod { name: "DrawEllipse",   arity: 6, target: MethodTarget::host("vybe:drawing", "drawEllipse") },
    DotnetMethod { name: "DrawArc",       arity: 8, target: MethodTarget::host("vybe:drawing", "drawArc") },
    DotnetMethod { name: "DrawString",    arity: 6, target: MethodTarget::host("vybe:drawing", "drawString") },
    DotnetMethod { name: "DrawImage",     arity: 4, target: MethodTarget::host("vybe:drawing", "drawImage") },
    // Brush-based fills.
    DotnetMethod { name: "FillRectangle", arity: 6, target: MethodTarget::host("vybe:drawing", "fillRectangle") },
    DotnetMethod { name: "FillEllipse",   arity: 6, target: MethodTarget::host("vybe:drawing", "fillEllipse") },
    DotnetMethod { name: "FillPolygon",   arity: 3, target: MethodTarget::host("vybe:drawing", "fillPolygon") },
    // Misc.
    DotnetMethod { name: "Clear",         arity: 2, target: MethodTarget::host("vybe:drawing", "clear") },
    DotnetMethod { name: "Dispose",       arity: 1, target: MethodTarget::host("vybe:drawing", "graphicsDispose") },
];

/// Methods on `Pen` and `Brush`. Both expose `Dispose` (no-op for now —
/// real .NET frees the underlying GDI handle).
const PEN_METHODS: &[DotnetMethod] = &[
    DotnetMethod { name: "Dispose", arity: 1, target: MethodTarget::host("vybe:drawing", "penDispose") },
];
const BRUSH_METHODS: &[DotnetMethod] = &[
    DotnetMethod { name: "Dispose", arity: 1, target: MethodTarget::host("vybe:drawing", "brushDispose") },
];

pub fn classes() -> &'static [DotnetClass] {
    &[
        // ── Graphics ────────────────────────────────────────────────────────
        // Returned by `Control.CreateGraphics()`. The user can also `New
        // Graphics(...)` — but real .NET doesn't expose a public ctor; we
        // expose one anyway for symmetry, and the host fn produces a fresh
        // graphics object regardless of args.
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
            widget_host_fn: Some("graphicsNew"),
            widget_host_module: "vybe:drawing",
        },
        // ── Pen ────────────────────────────────────────────────────────────
        // `New Pen(color, width)` — arity-2 constructor. Args are forwarded
        // to `vybe:drawing::penNew(color, width)`.
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
            widget_host_module: "vybe:drawing",
        },
        // ── Brush (abstract) ────────────────────────────────────────────────
        DotnetClass {
            name: "Brush",
            parent: Some("MarshalByRefObject"),
            properties: &[],
            methods: BRUSH_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:drawing",
        },
        // ── SolidBrush ──────────────────────────────────────────────────────
        // `New SolidBrush(color)` — arity-1 constructor.
        DotnetClass {
            name: "SolidBrush",
            parent: Some("Brush"),
            properties: &["Color"],
            methods: &[],
            ctor_arity: 1,
            widget_host_fn: Some("solidBrushNew"),
            widget_host_module: "vybe:drawing",
        },
        // ── HatchBrush ──────────────────────────────────────────────────────
        // Real .NET puts this in `System.Drawing.Drawing2D`. The user-side
        // shape is the same.
        DotnetClass {
            name: "HatchBrush",
            parent: Some("Brush"),
            properties: &["BackgroundColor", "ForegroundColor", "HatchStyle"],
            methods: &[],
            ctor_arity: 3,
            widget_host_fn: Some("hatchBrushNew"),
            widget_host_module: "vybe:drawing",
        },
        // ── LinearGradientBrush ─────────────────────────────────────────────
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
            widget_host_module: "vybe:drawing",
        },
    ]
}
