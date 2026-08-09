//! WHATWG HTML — `CanvasRenderingContext2D`.
//!
//! The 2D drawing surface, under its own spec names: `getContext`,
//! `fillRect`, `fillText`, `arc`, `beginPath`, `setLineDash`, `drawImage`.
//! A guest that knows the browser needs no translation, and a browser host
//! satisfies these imports with the real canvas element.
//!
//! Nothing here paints. Ops go to the engine installed via
//! [`crate::canvas_backend::set_backend`], which is what makes the renderer
//! swappable — native widgets today, a real browser engine later — without
//! the API moving.
//!
//! Adapters that speak another vocabulary live on THEIR side: SDL's
//! `SDL_FillRect` is `fillRect` plus its rect struct, `SDL_BlitPaletted` is
//! `drawImage` over paletted pixels, .NET's `Graphics`/`Pen`/`Brush` objects
//! are a dotnet-facing shim over `fillStyle`/`strokeStyle`/`lineWidth`.

use std::sync::Arc;

use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

use crate::canvas_backend::{Op2D, apply, backend};

/// `getContext` returns a handle; every op reads the target name back out of
/// it. Accepts the handle or a bare name, the way `getContext` itself accepts
/// an element or an id.
fn target_of(arg: Option<&Value>) -> String {
    match arg {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            o.properties
                .get("__control_name")
                .map(|v| format!("{}", v).to_lowercase())
                .unwrap_or_default()
        }
        Some(Value::Null) | Some(Value::Undefined) | None => String::new(),
        Some(other) => format!("{}", other).to_lowercase(),
    }
}

fn f32_arg(args: &[Value], idx: usize) -> f32 {
    args.get(idx).map(|v| v.as_f64() as f32).unwrap_or(0.0)
}

fn u8_arg(args: &[Value], idx: usize) -> u8 {
    args.get(idx)
        .map(|v| v.as_i32().clamp(0, 255) as u8)
        .unwrap_or(0)
}

fn bool_arg(args: &[Value], idx: usize) -> bool {
    args.get(idx).map(|v| v.as_bool()).unwrap_or(false)
}

/// Text argument decode. A guest may hand over a JS string, a C `char`
/// array holding codes, or a pointer view of one; all three are text.
fn text_arg(args: &[Value], idx: usize) -> String {
    fn from_value(v: &Value) -> String {
        match v {
            Value::String(s) => match s.find('\0') {
                Some(k) => s[..k].to_string(),
                None => s.to_string(),
            },
            Value::Object(obj) => {
                enum Shape {
                    Chars(String),
                    View(Option<Value>, usize),
                    Other,
                }
                let shape = {
                    let o = obj.lock().unwrap();
                    if let ObjectKind::Array(items) = &o.kind {
                        let mut out = String::new();
                        'items: for item in items.iter() {
                            match item {
                                Value::String(ch) => {
                                    if ch.is_empty() || ch.as_ref() == "\0" {
                                        break 'items;
                                    }
                                    out.push_str(ch);
                                }
                                Value::I32(0) => break 'items,
                                Value::I32(code) => {
                                    if let Some(c) = char::from_u32(*code as u32) {
                                        out.push(c);
                                    }
                                }
                                Value::F64(f) => {
                                    let code = *f as i64;
                                    if code == 0 {
                                        break 'items;
                                    }
                                    if let Some(c) = char::from_u32(code as u32) {
                                        out.push(c);
                                    }
                                }
                                Value::Null | Value::Undefined => break 'items,
                                other => out.push_str(&format!("{}", other)),
                            }
                        }
                        Shape::Chars(out)
                    } else if o
                        .properties
                        .get("__ref_kind")
                        .map(|k| format!("{}", k) == "carray")
                        .unwrap_or(false)
                    {
                        Shape::View(
                            o.properties.get("__base").cloned(),
                            o.properties
                                .get("__idx")
                                .map(|v| v.as_f64() as usize)
                                .unwrap_or(0),
                        )
                    } else {
                        Shape::Other
                    }
                };
                match shape {
                    Shape::Chars(s) => s,
                    Shape::View(Some(base), skip) => from_value(&base).chars().skip(skip).collect(),
                    Shape::View(None, _) | Shape::Other => String::new(),
                }
            }
            other => format!("{}", other),
        }
    }
    args.get(idx).map(from_value).unwrap_or_default()
}

/// Dense byte decode for `drawImage` pixel data.
fn bytes_arg(args: &[Value], idx: usize) -> Vec<u8> {
    let Some(v) = args.get(idx) else {
        return Vec::new();
    };
    let unwrapped = match v {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            if o.properties
                .get("__ref_kind")
                .map(|k| format!("{}", k) == "carray")
                .unwrap_or(false)
            {
                o.properties.get("__base").cloned()
            } else {
                None
            }
        }
        _ => None,
    };
    let target = unwrapped.unwrap_or_else(|| v.clone());
    match &target {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::Array(items) => items
                    .iter()
                    .map(|it| it.as_i32().clamp(0, 255) as u8)
                    .collect(),
                _ => Vec::new(),
            }
        }
        _ => Vec::new(),
    }
}

pub fn register(vm: &mut VM) {
    // ── getContext(target) ───────────────────────────────────────────────
    //
    // `element.getContext('2d')`. Returns a handle stamped with the target
    // name; framework wrappers (.NET `CreateGraphics`, Flutter's canvas
    // bridge) re-stamp `__type` with their own tag so guest code can
    // downcast, which is why the name travels in the object.
    vm.register_host_fn(
        "web:canvas",
        "getContext",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let name = args
                .first()
                .map(|v| format!("{}", v))
                .unwrap_or_default()
                .to_lowercase();
            if let Some(b) = backend() {
                b.ensure(&name);
            }
            let mut o = Object::new();
            o.properties.insert(
                "__type".into(),
                Value::String(Arc::from("CanvasRenderingContext2D")),
            );
            o.properties.insert(
                "__control_name".into(),
                Value::String(Arc::from(name.as_str())),
            );
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // ── state ────────────────────────────────────────────────────────────
    macro_rules! simple {
        ($name:literal, $build:expr) => {
            vm.register_host_fn(
                "web:canvas",
                $name,
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let target = target_of(args.first());
                    let build: fn(&[Value]) -> Op2D = $build;
                    apply(&target, build(args));
                    Value::Null
                }),
            );
        };
    }

    simple!("save", |_a| Op2D::Save);
    simple!("restore", |_a| Op2D::Restore);
    simple!("setFillStyle", |a| Op2D::SetFillStyle(
        u8_arg(a, 1),
        u8_arg(a, 2),
        u8_arg(a, 3),
        if a.len() > 4 { u8_arg(a, 4) } else { 255 }
    ));
    simple!("setStrokeStyle", |a| Op2D::SetStrokeStyle(
        u8_arg(a, 1),
        u8_arg(a, 2),
        u8_arg(a, 3),
        if a.len() > 4 { u8_arg(a, 4) } else { 255 }
    ));
    simple!("setLineWidth", |a| Op2D::SetLineWidth(f32_arg(a, 1)));
    simple!("setLineCap", |a| Op2D::SetLineCap(
        a.get(1).map(|v| format!("{}", v)).unwrap_or_default()
    ));
    simple!("setLineJoin", |a| Op2D::SetLineJoin(
        a.get(1).map(|v| format!("{}", v)).unwrap_or_default()
    ));
    simple!("setGlobalAlpha", |a| Op2D::SetGlobalAlpha(f32_arg(a, 1)));
    simple!("setImageSmoothingEnabled", |a| Op2D::SetImageSmoothing(
        bool_arg(a, 1)
    ));
    simple!("translate", |a| Op2D::Translate(
        f32_arg(a, 1),
        f32_arg(a, 2)
    ));
    simple!("scale", |a| Op2D::Scale(f32_arg(a, 1), f32_arg(a, 2)));
    simple!("rotate", |a| Op2D::Rotate(f32_arg(a, 1)));

    // `setLineDash([...])` — the spec takes a sequence; the dash lengths
    // arrive as trailing numeric args, empty meaning solid.
    vm.register_host_fn(
        "web:canvas",
        "setLineDash",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            let dashes: Vec<f32> = args
                .iter()
                .skip(1)
                .map(|v| v.as_f64() as f32)
                .filter(|d| *d > 0.0)
                .collect();
            apply(&target, Op2D::SetLineDash(dashes));
            Value::Null
        }),
    );

    // `font = "italic bold 16px Arial"`, pre-parsed by the caller.
    vm.register_host_fn(
        "web:canvas",
        "setFont",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            apply(
                &target,
                Op2D::SetFont {
                    family: args.get(1).map(|v| format!("{}", v)).unwrap_or_default(),
                    size: f32_arg(args, 2),
                    bold: bool_arg(args, 3),
                    italic: bool_arg(args, 4),
                },
            );
            Value::Null
        }),
    );

    // ── paths ────────────────────────────────────────────────────────────
    simple!("beginPath", |_a| Op2D::BeginPath);
    simple!("closePath", |_a| Op2D::ClosePath);
    simple!("moveTo", |a| Op2D::MoveTo(f32_arg(a, 1), f32_arg(a, 2)));
    simple!("lineTo", |a| Op2D::LineTo(f32_arg(a, 1), f32_arg(a, 2)));
    simple!("arc", |a| Op2D::Arc(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4),
        f32_arg(a, 5),
        bool_arg(a, 6)
    ));
    simple!("bezierCurveTo", |a| Op2D::BezierCurveTo(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4),
        f32_arg(a, 5),
        f32_arg(a, 6)
    ));
    simple!("quadraticCurveTo", |a| Op2D::QuadraticCurveTo(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ));
    simple!("rect", |a| Op2D::Rect(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ));
    simple!("fill", |_a| Op2D::Fill);
    simple!("stroke", |_a| Op2D::Stroke);
    simple!("clip", |_a| Op2D::Clip);

    // ── shapes / text ────────────────────────────────────────────────────
    simple!("fillRect", |a| Op2D::FillRect(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ));
    simple!("strokeRect", |a| Op2D::StrokeRect(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ));
    simple!("clearRect", |a| Op2D::ClearRect(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ));

    vm.register_host_fn(
        "web:canvas",
        "fillText",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            apply(
                &target,
                Op2D::FillText(text_arg(args, 1), f32_arg(args, 2), f32_arg(args, 3)),
            );
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:canvas",
        "strokeText",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            apply(
                &target,
                Op2D::StrokeText(text_arg(args, 1), f32_arg(args, 2), f32_arg(args, 3)),
            );
            Value::Null
        }),
    );

    // ── images ───────────────────────────────────────────────────────────
    //
    // `drawImage(image, dx, dy, dw, dh)` where the image is dense RGBA —
    // `putImageData`'s territory, and the frame path of a software renderer.
    vm.register_host_fn(
        "web:canvas",
        "drawImage",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            apply(
                &target,
                Op2D::DrawImageRgba {
                    pixels: bytes_arg(args, 1),
                    width: args.get(2).map(|v| v.as_i32().max(0) as u32).unwrap_or(0),
                    height: args.get(3).map(|v| v.as_i32().max(0) as u32).unwrap_or(0),
                    dx: f32_arg(args, 4),
                    dy: f32_arg(args, 5),
                    dw: f32_arg(args, 6),
                    dh: f32_arg(args, 7),
                },
            );
            Value::Null
        }),
    );

    // Paletted variant: 8-bit indices + a 256-entry RGB palette, expanded by
    // the backend. Not a DOM method — an extension for palette-era software
    // renderers, kept here because it is a canvas concern, not an SDL one.
    vm.register_host_fn(
        "web:canvas",
        "drawImagePaletted",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            apply(
                &target,
                Op2D::DrawImagePaletted {
                    indices: bytes_arg(args, 1),
                    palette: bytes_arg(args, 2),
                    width: args.get(3).map(|v| v.as_i32().max(0) as u32).unwrap_or(0),
                    height: args.get(4).map(|v| v.as_i32().max(0) as u32).unwrap_or(0),
                    dx: f32_arg(args, 5),
                    dy: f32_arg(args, 6),
                    dw: f32_arg(args, 7),
                    dh: f32_arg(args, 8),
                },
            );
            Value::Null
        }),
    );

    // Drop every recorded op for a target — the frame-start reset a
    // double-buffered renderer needs.
    vm.register_host_fn(
        "web:canvas",
        "clearAll",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            if let Some(b) = backend() {
                b.clear_all(&target);
            }
            Value::Null
        }),
    );
}
