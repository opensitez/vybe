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

/// The surface a context handle draws on.
///
/// A context is bound to the ELEMENT it came from — `canvas.getContext('2d')`
/// in every browser — so the handle carries `__node` and the target is derived
/// from it. The backend's target is a string because a backend is below the
/// seam and may key its surfaces however it likes; what must be spec-shaped is
/// the guest-facing API, and that is an element.
///
/// The bare-name form is a **migration path, not the API**. .NET
/// `CreateGraphics` and Flutter's canvas bridge pass a control name today
/// because that is what the retired host took. A real browser engine has no control
/// name to resolve, so anything that depends on this cannot survive an engine
/// swap — see the `__control_name` note on `get_context`.
fn target_of(arg: Option<&Value>) -> String {
    match arg {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            // An element-bound context: the node IS the identity.
            if let Some(node) = o.properties.get("__node") {
                return format!("n{}", node.as_f64() as u64);
            }
            o.properties
                .get("__control_name")
                .map(|v| format!("{}", v).to_lowercase())
                .unwrap_or_default()
        }
        Some(Value::Null) | Some(Value::Undefined) | None => String::new(),
        Some(other) => format!("{}", other).to_lowercase(),
    }
}

/// Read an element handle's node id, if the argument is one.
fn node_of(arg: Option<&Value>) -> Option<u64> {
    match arg {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            o.properties.get("__node").map(|v| v.as_f64() as u64)
        }
        _ => None,
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
    // ── getContext(element, contextType) ─────────────────────────────────
    //
    // **`canvas.getContext('2d')`** — HTML §4.12.5. A context belongs to an
    // ELEMENT and is asked for by context type; that is the whole signature,
    // and it is the one a browser engine can implement.
    //
    // It used to take a lower-cased control NAME as argument 0 and ignore the
    // context type entirely, while this comment claimed otherwise. A real
    // engine has no control name to resolve, so `web:canvas` could not have
    // survived the swap the seam exists for — the API worked and was not
    // implementable, which is the failure mode that hides longest.
    //
    // The name form still resolves, because .NET `CreateGraphics` and
    // Flutter's canvas bridge pass one today. It is a MIGRATION PATH: a caller
    // that made the element itself already holds the handle and needs no
    // lookup at all. Framework wrappers re-stamp `__type` with their own tag so
    // guest code can downcast, which is why the identity travels in the object.
    vm.register_host_fn(
        "web:canvas",
        "getContext",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let node = node_of(args.first());
            // `getContext` answers null for a context type it does not
            // support, per spec — not a handle that silently paints nowhere.
            // Only checked on the element form: the legacy name form passed the
            // name here and has no type argument to read.
            if node.is_some() {
                let context_type = args
                    .get(1)
                    .map(|v| format!("{}", v).to_lowercase())
                    .unwrap_or_default();
                if context_type != "2d" {
                    return Value::Null;
                }
            }
            let target = match node {
                Some(id) => format!("n{id}"),
                None => args
                    .first()
                    .map(|v| format!("{}", v))
                    .unwrap_or_default()
                    .to_lowercase(),
            };
            if let Some(b) = backend() {
                b.ensure(&target);
            }
            let mut o = Object::new();
            o.properties.insert(
                "__type".into(),
                Value::String(Arc::from("CanvasRenderingContext2D")),
            );
            match node {
                // Element-bound: the node is the identity, and `target_of`
                // derives the surface from it.
                Some(id) => {
                    o.properties.insert("__node".into(), Value::F64(id as f64));
                }
                None => {
                    o.properties.insert(
                        "__control_name".into(),
                        Value::String(Arc::from(target.as_str())),
                    );
                }
            }
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
    // The rest of the 2D context's state and path surface. Every one of these
    // was already implemented by the engine and unreachable from a page,
    // because `Op2D` — the wire format between the API and the painter — had no
    // variant to carry it. Adding the variants was blocked until `platforms/vybe`
    // stopped implementing `CanvasBackend`, which it now does not.
    simple!("transform", |a| Op2D::Transform(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4),
        f32_arg(a, 5),
        f32_arg(a, 6)
    ));
    simple!("resetTransform", |_a| Op2D::ResetTransform);
    // `setTransform(a, b, c, d, e, f)` REPLACES the current matrix, where
    // `transform` multiplies into it. Emitted as reset-then-transform rather
    // than as a third op, so the painter has one notion of "apply a matrix" and
    // the difference stays where it belongs — in what the API asked for.
    vm.register_host_fn(
        "web:canvas",
        "setTransform",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            apply(&target, Op2D::ResetTransform);
            apply(
                &target,
                Op2D::Transform(
                    f32_arg(args, 1),
                    f32_arg(args, 2),
                    f32_arg(args, 3),
                    f32_arg(args, 4),
                    f32_arg(args, 5),
                    f32_arg(args, 6),
                ),
            );
            Value::Null
        }),
    );
    simple!("setMiterLimit", |a| Op2D::SetMiterLimit(f32_arg(a, 1)));
    simple!("setTextAlign", |a| Op2D::SetTextAlign(
        a.get(1).map(|v| format!("{}", v)).unwrap_or_default()
    ));
    simple!("setTextBaseline", |a| Op2D::SetTextBaseline(
        a.get(1).map(|v| format!("{}", v)).unwrap_or_default()
    ));
    simple!("setLineDashOffset", |a| Op2D::SetLineDashOffset(f32_arg(a, 1)));
    // `ellipse(x, y, radiusX, radiusY, rotation, startAngle, endAngle, ccw)` —
    // the trailing five are accepted and dropped: the engine draws an
    // axis-aligned full ellipse. Taking them and ignoring them is honest about
    // the call site's shape; refusing the extra arguments would make a
    // spec-correct call fail.
    simple!("ellipse", |a| Op2D::Ellipse(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ));

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

    // ── ImageData ────────────────────────────────────────────────────────
    //
    // `createImageData(sw, sh)` and `putImageData(imagedata, dx, dy)` — the
    // spec's route for handing raw pixels to a canvas, and the ONLY one.
    //
    // `drawImage` takes a `CanvasImageSource` — an `HTMLImageElement`, an
    // `ImageBitmap`, another canvas — and never a byte array. So a software
    // renderer that has computed a frame and holds no element has exactly this
    // door, and the byte-array `drawImage` below is not it: no browser has
    // that signature, which means nothing depending on it could survive an
    // engine swap.
    //
    // `ImageData` is a plain object with `data`, `width`, `height`. `data` is
    // a `Uint8ClampedArray` in a browser; here it is the array the guest
    // already had, because the guest is what fills it.
    vm.register_host_fn(
        "web:canvas",
        "createImageData",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            // `createImageData(sw, sh)` on a context — arg 0 is the context,
            // which this does not need: an `ImageData` is not bound to a
            // canvas. It is taken so the call shape matches every other
            // method here and the receiver form works.
            let width = args.get(1).map(|v| v.as_i32().max(0)).unwrap_or(0);
            let height = args.get(2).map(|v| v.as_i32().max(0)).unwrap_or(0);
            // Transparent black, per spec — every byte zero, INCLUDING alpha.
            let bytes = (width as usize).saturating_mul(height as usize) * 4;
            let data = Object::new_array(vec![Value::I32(0); bytes]);
            let mut image_data = Object::new();
            image_data
                .properties
                .insert("__type".into(), Value::String("ImageData".into()));
            image_data.properties.insert(
                "data".into(),
                Value::Object(vybe_runtime::heap::alloc(data)),
            );
            image_data
                .properties
                .insert("width".into(), Value::I32(width));
            image_data
                .properties
                .insert("height".into(), Value::I32(height));
            Value::Object(vybe_runtime::heap::alloc(image_data))
        }),
    );
    vm.register_host_fn(
        "web:canvas",
        "putImageData",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            // The dimensions come from the `ImageData`, not from the caller —
            // that is what makes it an ImageData rather than three loose
            // arguments, and it is why `putImageData` cannot disagree with
            // itself about the shape of the buffer.
            let (pixels, width, height) = match args.get(1) {
                Some(Value::Object(o)) => {
                    let bag = o.lock().unwrap();
                    let width = bag.properties.get("width").map(|v| v.as_f64() as u32);
                    let height = bag.properties.get("height").map(|v| v.as_f64() as u32);
                    let data = bag.properties.get("data").cloned();
                    drop(bag);
                    match (data, width, height) {
                        (Some(data), Some(w), Some(h)) => (bytes_arg(&[data], 0), w, h),
                        _ => (Vec::new(), 0, 0),
                    }
                }
                _ => (Vec::new(), 0, 0),
            };
            // A buffer that does not match its own dimensions is not a
            // partial write, it is a caller mistake — dropped rather than
            // painted from whatever bytes happened to be there.
            if pixels.len() != (width as usize) * (height as usize) * 4 {
                return Value::Null;
            }
            apply(
                &target,
                Op2D::PutImageData {
                    pixels,
                    width,
                    height,
                    dx: f32_arg(args, 2),
                    dy: f32_arg(args, 3),
                },
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

    // `context.measureText(text)` — HTML §4.12.5.
    //
    // The only query in the whole 2D context, and the reason the seam grew a
    // return path: every other call paints and answers nothing.
    //
    // Returns a `TextMetrics` carrying `width` and NOTHING else. The spec has
    // five more members — `actualBoundingBoxAscent`/`Descent`,
    // `fontBoundingBoxAscent`/`Descent`, `emHeightAscent` — and the engine
    // measures a line box, not a glyph box, so it cannot answer them. They are
    // ABSENT rather than filled with a plausible number derived from the
    // height: a synthesised ascent reads as measured, and a caller laying text
    // out against it would be wrong in a way nothing could see. `width` is
    // also the one member every engine has always had.
    //
    // `null` when the target names no surface — distinguishable from a zero
    // width, which is a legal answer for the empty string.
    vm.register_host_fn(
        "web:canvas",
        "measureText",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            let text = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => String::new(),
            };
            let Some(width) = backend().and_then(|b| b.measure_text(&target, &text)) else {
                return Value::Null;
            };
            let mut metrics = Object::new();
            metrics
                .properties
                .insert("__type".into(), Value::String("TextMetrics".into()));
            metrics
                .properties
                .insert("width".into(), Value::F64(width as f64));
            Value::Object(vybe_runtime::heap::alloc(metrics))
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
