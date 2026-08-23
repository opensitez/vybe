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

use std::sync::{Arc, OnceLock};

use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

use crate::canvas_backend::{
    GradientDef, GradientKind, Op2D, PathDef, PathOp2D, PatternDef, Query2D, Query2DValue,
    StringAttribute, apply, backend, query,
};

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

/// An argument as a string.
///
/// `format!` rather than a `Value::String` match, because a guest may hand over
/// a number where the IDL says `DOMString` — `ctx.font = 16` is not valid CSS,
/// but stringifying it and letting the engine reject it is what a browser does,
/// and it keeps the failure in one place.
fn str_arg(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}

/// The pixels of an `ImageData`-shaped argument: `{data, width, height}`.
///
/// `None` when the object is not that shape or its buffer does not match its
/// own dimensions — which is a caller mistake, not a partial image, and is
/// dropped rather than read from whatever bytes happened to be there.
fn image_data_arg(arg: Option<&Value>) -> Option<(Vec<u8>, u32, u32)> {
    let Some(Value::Object(o)) = arg else {
        return None;
    };
    let bag = o.lock().unwrap();
    let width = bag.properties.get("width").map(|v| v.as_f64() as u32)?;
    let height = bag.properties.get("height").map(|v| v.as_f64() as u32)?;
    let data = bag.properties.get("data").cloned()?;
    drop(bag);
    let pixels = bytes_arg(&[data], 0);
    if pixels.len() != (width as usize) * (height as usize) * 4 {
        return None;
    }
    Some((pixels, width, height))
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


/// The IDL ATTRIBUTES of `CanvasRenderingContext2D`, and the host functions
/// that read and write each one.
///
/// **This is what makes `ctx.fillStyle = "red"` work.** Everything in the
/// interface that is not a method is one of these, and until they were wired a
/// page could only reach them by calling `setFillStyle(...)` — a spelling no
/// page has, invented because the seam had no other shape.
///
/// `(attribute, getter, setter)`. A `""` setter is a read-only attribute.
const ATTRIBUTES: &[(&str, &str, &str)] = &[
    ("fillStyle", "getFillStyle", "setFillStyleCss"),
    ("strokeStyle", "getStrokeStyle", "setStrokeStyleCss"),
    ("font", "getFont", "setFontCss"),
    ("filter", "getFilter", "setFilter"),
    ("globalAlpha", "", "setGlobalAlpha"),
    ("lineWidth", "", "setLineWidth"),
    ("lineCap", "getLineCap", "setLineCap"),
    ("lineJoin", "getLineJoin", "setLineJoin"),
    ("miterLimit", "", "setMiterLimit"),
    ("lineDashOffset", "", "setLineDashOffset"),
    ("textAlign", "getTextAlign", "setTextAlign"),
    ("textBaseline", "getTextBaseline", "setTextBaseline"),
    ("direction", "getDirection", "setDirection"),
    ("letterSpacing", "getLetterSpacing", "setLetterSpacing"),
    ("wordSpacing", "getWordSpacing", "setWordSpacing"),
    ("fontKerning", "getFontKerning", "setFontKerning"),
    ("fontStretch", "getFontStretch", "setFontStretch"),
    ("fontVariantCaps", "getFontVariantCaps", "setFontVariantCaps"),
    ("textRendering", "getTextRendering", "setTextRendering"),
    ("lang", "getLang", "setLang"),
    ("shadowColor", "getShadowColor", "setShadowColor"),
    ("shadowBlur", "", "setShadowBlur"),
    ("shadowOffsetX", "", "setShadowOffsetX"),
    ("shadowOffsetY", "", "setShadowOffsetY"),
    (
        "globalCompositeOperation",
        "getGlobalCompositeOperation",
        "setGlobalCompositeOperation",
    ),
    ("imageSmoothingEnabled", "", "setImageSmoothingEnabled"),
    (
        "imageSmoothingQuality",
        "getImageSmoothingQuality",
        "setImageSmoothingQuality",
    ),
];

/// Resolved indices for [`ATTRIBUTES`], filled once the host functions are
/// registered.
static ATTR_FNS: OnceLock<Vec<(&'static str, Option<usize>, Option<usize>)>> = OnceLock::new();

/// A callable reference to a host function, for an accessor slot.
fn host_fn_ref(name: &str, idx: usize) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__host_module".into(), Value::String("web:canvas".into()));
    obj.properties
        .insert("__host_name".into(), Value::String(name.into()));
    obj.properties
        .insert("__host_idx".into(), Value::F64(idx as f64));
    obj.properties
        .insert("name".into(), Value::String(name.into()));
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// Hang `__get_<attr>` / `__set_<attr>` on a freshly made context.
///
/// The VM's own accessor protocol (ECMA-262 §10.1.8.1, `dispatch.rs`): a
/// property read finds `__get_<name>` and calls it with the receiver, and a
/// write finds `__set_<name>` and calls it with the receiver and the value.
/// Both are exactly the shape the `web:canvas` host functions already have,
/// because a vtable-dispatched host fn takes its receiver as argument 0 — so
/// this wires the spec spelling to the functions that were already there
/// rather than adding a second set.
fn install_accessors(ctx: &mut Object) {
    let Some(table) = ATTR_FNS.get() else { return };
    for (attr, getter, setter) in table {
        if let Some(idx) = getter {
            ctx.properties
                .insert(format!("__get_{attr}").into(), host_fn_ref(attr, *idx));
        }
        if let Some(idx) = setter {
            ctx.properties
                .insert(format!("__set_{attr}").into(), host_fn_ref(attr, *idx));
        }
    }
}


/// What `fillStyle` / `strokeStyle` was assigned.
enum StyleValue {
    Css(String),
    Gradient(GradientDef),
    Pattern(PatternDef),
}

/// Classify an assignment to `fillStyle` / `strokeStyle`.
///
/// The IDL type is `DOMString | CanvasGradient | CanvasPattern`, and the three
/// are told apart by the `__type` the factories stamp. Anything else is treated
/// as a CSS string, which is what an unrecognised value is.
fn style_value(v: Option<&Value>) -> StyleValue {
    let Some(Value::Object(o)) = v else {
        return StyleValue::Css(v.map(|x| format!("{}", x)).unwrap_or_default());
    };
    let lock = o.lock().unwrap();
    let kind = lock
        .properties
        .get("__type")
        .map(|t| format!("{}", t))
        .unwrap_or_default();
    match kind.as_str() {
        "CanvasGradient" => {
            let coords = lock
                .properties
                .get("__coords")
                .map(|c| numbers(c))
                .unwrap_or_default();
            let at = |i: usize| coords.get(i).copied().unwrap_or(0.0);
            let kind = match lock
                .properties
                .get("__kind")
                .map(|k| format!("{}", k))
                .unwrap_or_default()
                .as_str()
            {
                "radial" => GradientKind::Radial {
                    x0: at(0),
                    y0: at(1),
                    r0: at(2),
                    x1: at(3),
                    y1: at(4),
                    r1: at(5),
                },
                "conic" => GradientKind::Conic {
                    angle: at(0),
                    x: at(1),
                    y: at(2),
                },
                _ => GradientKind::Linear {
                    x0: at(0),
                    y0: at(1),
                    x1: at(2),
                    y1: at(3),
                },
            };
            let mut stops = Vec::new();
            if let Some(Value::Object(list)) = lock.properties.get("__stops") {
                let list = list.lock().unwrap();
                if let ObjectKind::Array(ref items) = list.kind {
                    for item in items {
                        if let Value::Object(stop) = item {
                            let stop = stop.lock().unwrap();
                            let offset = stop
                                .properties
                                .get("offset")
                                .map(|v| v.as_f64() as f32)
                                .unwrap_or(0.0);
                            let color = stop
                                .properties
                                .get("color")
                                .map(|v| format!("{}", v))
                                .unwrap_or_default();
                            stops.push((offset, color));
                        }
                    }
                }
            }
            StyleValue::Gradient(GradientDef { kind, stops })
        }
        "CanvasPattern" => {
            let width = lock
                .properties
                .get("__width")
                .map(|v| v.as_f64() as u32)
                .unwrap_or(0);
            let height = lock
                .properties
                .get("__height")
                .map(|v| v.as_f64() as u32)
                .unwrap_or(0);
            let pixels = lock
                .properties
                .get("__pixels")
                .map(|p| bytes_arg(&[p.clone()], 0))
                .unwrap_or_default();
            let repetition = lock
                .properties
                .get("__repetition")
                .map(|v| format!("{}", v))
                .unwrap_or_else(|| "repeat".to_string());
            StyleValue::Pattern(PatternDef {
                pixels,
                width,
                height,
                repetition,
            })
        }
        // A plain object is not a style; stringifying it gives the engine
        // something it will reject, which leaves the attribute alone — the
        // spec's rule for an unparseable value.
        _ => StyleValue::Css(kind),
    }
}

/// The numbers in an array-valued property.
fn numbers(v: &Value) -> Vec<f32> {
    let Value::Object(o) = v else {
        return Vec::new();
    };
    let lock = o.lock().unwrap();
    match lock.kind {
        ObjectKind::Array(ref items) => items.iter().map(|x| x.as_f64() as f32).collect(),
        _ => Vec::new(),
    }
}


/// Read a `Path2D` argument back into the seam's form.
///
/// A `Path2D` is a plain object carrying `__ops`, appended to by its own
/// methods. `None` when the value is not one — which is how the overloaded
/// `fill` / `stroke` / `clip` tell `fill(path)` from `fill("evenodd")`.
fn path_arg(v: Option<&Value>) -> Option<PathDef> {
    let Some(Value::Object(o)) = v else {
        return None;
    };
    let lock = o.lock().unwrap();
    if lock.properties.get("__type").map(|t| format!("{}", t)).as_deref() != Some("Path2D") {
        return None;
    }
    let Some(Value::Object(list)) = lock.properties.get("__ops") else {
        return Some(PathDef::default());
    };
    let list = list.lock().unwrap();
    let ObjectKind::Array(ref items) = list.kind else {
        return Some(PathDef::default());
    };
    let mut ops = Vec::with_capacity(items.len());
    for item in items {
        let Value::Object(seg) = item else { continue };
        let seg = seg.lock().unwrap();
        let n = |k: &str| {
            seg.properties
                .get(k)
                .map(|v| v.as_f64() as f32)
                .unwrap_or(0.0)
        };
        let flag = |k: &str| matches!(seg.properties.get(k), Some(Value::Bool(true)));
        let kind = seg
            .properties
            .get("op")
            .map(|v| format!("{}", v))
            .unwrap_or_default();
        ops.push(match kind.as_str() {
            "closePath" => PathOp2D::ClosePath,
            "moveTo" => PathOp2D::MoveTo(n("x"), n("y")),
            "lineTo" => PathOp2D::LineTo(n("x"), n("y")),
            "quadraticCurveTo" => PathOp2D::QuadraticCurveTo {
                cx: n("cx"),
                cy: n("cy"),
                x: n("x"),
                y: n("y"),
            },
            "bezierCurveTo" => PathOp2D::BezierCurveTo {
                cx1: n("cx1"),
                cy1: n("cy1"),
                cx2: n("cx2"),
                cy2: n("cy2"),
                x: n("x"),
                y: n("y"),
            },
            "arcTo" => PathOp2D::ArcTo {
                x1: n("x1"),
                y1: n("y1"),
                x2: n("x2"),
                y2: n("y2"),
                radius: n("radius"),
            },
            "rect" => PathOp2D::Rect {
                x: n("x"),
                y: n("y"),
                w: n("w"),
                h: n("h"),
            },
            "roundRect" => PathOp2D::RoundRect {
                x: n("x"),
                y: n("y"),
                w: n("w"),
                h: n("h"),
                radii: [n("r0"), n("r1"), n("r2"), n("r3")],
            },
            "arc" => PathOp2D::Arc {
                x: n("x"),
                y: n("y"),
                r: n("r"),
                start: n("start"),
                end: n("end"),
                ccw: flag("ccw"),
            },
            "ellipse" => PathOp2D::Ellipse {
                x: n("x"),
                y: n("y"),
                rx: n("rx"),
                ry: n("ry"),
                rotation: n("rotation"),
                start: n("start"),
                end: n("end"),
                ccw: flag("ccw"),
            },
            // An unknown segment is DROPPED rather than guessed at: a wrong
            // segment draws a wrong shape, which is worse than a missing one.
            _ => continue,
        });
    }
    Some(PathDef { ops })
}

/// Append one already-parsed segment to a `Path2D` object.
///
/// The counterpart to `path_arg`: used by `addPath` and by the copy
/// constructor, which both take segments OUT of one path and put them into
/// another. Segments are copied rather than shared, so later edits to either
/// path leave the other alone — `addPath` is a snapshot, not a link.
fn copy_path_op(path: Option<&Value>, op: &PathOp2D) {
    let f = |v: f32| Value::F64(v as f64);
    match *op {
        PathOp2D::ClosePath => push_path_op(path, "closePath", &[]),
        PathOp2D::MoveTo(x, y) => push_path_op(path, "moveTo", &[("x", f(x)), ("y", f(y))]),
        PathOp2D::LineTo(x, y) => push_path_op(path, "lineTo", &[("x", f(x)), ("y", f(y))]),
        PathOp2D::QuadraticCurveTo { cx, cy, x, y } => push_path_op(
            path,
            "quadraticCurveTo",
            &[("cx", f(cx)), ("cy", f(cy)), ("x", f(x)), ("y", f(y))],
        ),
        PathOp2D::BezierCurveTo { cx1, cy1, cx2, cy2, x, y } => push_path_op(
            path,
            "bezierCurveTo",
            &[
                ("cx1", f(cx1)),
                ("cy1", f(cy1)),
                ("cx2", f(cx2)),
                ("cy2", f(cy2)),
                ("x", f(x)),
                ("y", f(y)),
            ],
        ),
        PathOp2D::ArcTo { x1, y1, x2, y2, radius } => push_path_op(
            path,
            "arcTo",
            &[
                ("x1", f(x1)),
                ("y1", f(y1)),
                ("x2", f(x2)),
                ("y2", f(y2)),
                ("radius", f(radius)),
            ],
        ),
        PathOp2D::Rect { x, y, w, h } => push_path_op(
            path,
            "rect",
            &[("x", f(x)), ("y", f(y)), ("w", f(w)), ("h", f(h))],
        ),
        PathOp2D::RoundRect { x, y, w, h, radii } => push_path_op(
            path,
            "roundRect",
            &[
                ("x", f(x)),
                ("y", f(y)),
                ("w", f(w)),
                ("h", f(h)),
                ("r0", f(radii[0])),
                ("r1", f(radii[1])),
                ("r2", f(radii[2])),
                ("r3", f(radii[3])),
            ],
        ),
        PathOp2D::Arc { x, y, r, start, end, ccw } => push_path_op(
            path,
            "arc",
            &[
                ("x", f(x)),
                ("y", f(y)),
                ("r", f(r)),
                ("start", f(start)),
                ("end", f(end)),
                ("ccw", Value::Bool(ccw)),
            ],
        ),
        PathOp2D::Ellipse { x, y, rx, ry, rotation, start, end, ccw } => push_path_op(
            path,
            "ellipse",
            &[
                ("x", f(x)),
                ("y", f(y)),
                ("rx", f(rx)),
                ("ry", f(ry)),
                ("rotation", f(rotation)),
                ("start", f(start)),
                ("end", f(end)),
                ("ccw", Value::Bool(ccw)),
            ],
        ),
    }
}

/// Append one segment to a `Path2D` object's `__ops`.
fn push_path_op(path: Option<&Value>, op: &str, fields: &[(&str, Value)]) {
    let Some(Value::Object(o)) = path else { return };
    let lock = o.lock().unwrap();
    let Some(Value::Object(list)) = lock.properties.get("__ops") else {
        return;
    };
    let mut list = list.lock().unwrap();
    let ObjectKind::Array(ref mut items) = list.kind else {
        return;
    };
    let mut seg = Object::new();
    seg.properties
        .insert("op".into(), Value::String(op.into()));
    for (k, v) in fields {
        seg.properties.insert((*k).into(), v.clone());
    }
    items.push(Value::Object(vybe_runtime::heap::alloc(seg)));
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
            // The IDL attributes. Without these `ctx.fillStyle = "red"` would
            // set a plain property on the object and paint nothing — the write
            // would succeed, which is the worst way for it to fail.
            install_accessors(&mut o);
            // `ctx.canvas` — the element this context belongs to. Read-only,
            // and a plain property because the answer never changes: a context
            // is bound to one element for its whole life. A page uses it to
            // reach `canvas.width` from a function that was handed only the
            // context, which is most functions that draw.
            if let Some(id) = node {
                o.properties.insert("canvas".into(), Value::F64(id as f64));
            }
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
            let Query2DValue::Metrics(m) = query(&target, Query2D::MeasureText(text)) else {
                return Value::Null;
            };
            let mut metrics = Object::new();
            metrics
                .properties
                .insert("__type".into(), Value::String("TextMetrics".into()));
            // **All twelve, not just `width`.** The other eleven used to be
            // absent — the engine computed them and the seam had nowhere to put
            // them, so this object claimed a canvas could not answer questions
            // it had already answered.
            for (name, value) in [
                ("width", m.width),
                ("actualBoundingBoxLeft", m.actual_bounding_box_left),
                ("actualBoundingBoxRight", m.actual_bounding_box_right),
                ("actualBoundingBoxAscent", m.actual_bounding_box_ascent),
                ("actualBoundingBoxDescent", m.actual_bounding_box_descent),
                ("fontBoundingBoxAscent", m.font_bounding_box_ascent),
                ("fontBoundingBoxDescent", m.font_bounding_box_descent),
                ("emHeightAscent", m.em_height_ascent),
                ("emHeightDescent", m.em_height_descent),
                ("hangingBaseline", m.hanging_baseline),
                ("alphabeticBaseline", m.alphabetic_baseline),
                ("ideographicBaseline", m.ideographic_baseline),
            ] {
                metrics
                    .properties
                    .insert(name.into(), Value::F64(value as f64));
            }
            Value::Object(vybe_runtime::heap::alloc(metrics))
        }),
    );

    // ── The rest of §4.12.5 ──────────────────────────────────────────────
    //
    // Everything below was implemented by BOTH engines and reachable from
    // nothing: the seam carried 40-odd operations and the interface has 75
    // members, so shadows, filters, gradients, the eight text-style attributes,
    // `arcTo`, `roundRect` and every query were dead ends. Registering them is
    // what makes "the same API on both engines" mean something to a page rather
    // than only to the two crates.

    // String-valued attributes. The VALUE crosses as the author wrote it and
    // the ENGINE parses it, with the parser its own stylesheets go through —
    // `platforms/web` deliberately parses no CSS.
    macro_rules! css_attr {
        ($name:literal, $build:expr) => {
            simple!($name, |a| {
                let build: fn(String) -> Op2D = $build;
                build(str_arg(a, 1))
            });
        };
    }
    // `fillStyle` / `strokeStyle` are `DOMString | CanvasGradient |
    // CanvasPattern`, so these cannot just stringify their argument: a gradient
    // formatted as text is not CSS, the engine would fail to parse it, and the
    // spec's own rule — an unparseable value is IGNORED — would turn assigning
    // a gradient into a silent no-op. Which of the three it is has to be
    // decided here, where the object is still an object.
    macro_rules! style_attr {
        ($name:literal, $css:expr, $grad:expr, $pat:expr) => {
            vm.register_host_fn(
                "web:canvas",
                $name,
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let target = target_of(args.first());
                    let value = args.get(1);
                    let op = match style_value(value) {
                        StyleValue::Gradient(def) => {
                            let build: fn(GradientDef) -> Op2D = $grad;
                            build(def)
                        }
                        StyleValue::Pattern(def) => {
                            let build: fn(PatternDef) -> Op2D = $pat;
                            build(def)
                        }
                        StyleValue::Css(css) => {
                            let build: fn(String) -> Op2D = $css;
                            build(css)
                        }
                    };
                    apply(&target, op);
                    Value::Null
                }),
            );
        };
    }
    style_attr!(
        "setFillStyleCss",
        Op2D::SetFillStyleCss,
        Op2D::SetFillGradient,
        Op2D::SetFillPattern
    );
    style_attr!(
        "setStrokeStyleCss",
        Op2D::SetStrokeStyleCss,
        Op2D::SetStrokeGradient,
        Op2D::SetStrokePattern
    );
    css_attr!("setFontCss", Op2D::SetFontCss);
    css_attr!("setFilter", Op2D::SetFilter);
    css_attr!("setShadowColor", Op2D::SetShadowColor);
    css_attr!(
        "setGlobalCompositeOperation",
        Op2D::SetGlobalCompositeOperation
    );
    css_attr!("setImageSmoothingQuality", Op2D::SetImageSmoothingQuality);
    css_attr!("setDirection", Op2D::SetDirection);
    css_attr!("setLetterSpacing", Op2D::SetLetterSpacing);
    css_attr!("setWordSpacing", Op2D::SetWordSpacing);
    css_attr!("setFontKerning", Op2D::SetFontKerning);
    css_attr!("setFontStretch", Op2D::SetFontStretch);
    css_attr!("setFontVariantCaps", Op2D::SetFontVariantCaps);
    css_attr!("setTextRendering", Op2D::SetTextRendering);
    css_attr!("setLang", Op2D::SetLang);

    // Shadows — three numbers and the colour above.
    simple!("setShadowBlur", |a| Op2D::SetShadowBlur(f32_arg(a, 1)));
    simple!("setShadowOffsetX", |a| Op2D::SetShadowOffsetX(f32_arg(a, 1)));
    simple!("setShadowOffsetY", |a| Op2D::SetShadowOffsetY(f32_arg(a, 1)));

    // Paths the seam could not express.
    simple!("arcTo", |a| Op2D::ArcTo(
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4),
        f32_arg(a, 5)
    ));
    // `roundRect(x, y, w, h, radii)` — the IDL takes one radius, a list of up
    // to four, or a `DOMPointInit`. The short forms are expanded HERE because
    // that is a signature rule, not a rendering one: one radius means all four
    // corners, two means the diagonals, three means top-left / both others /
    // bottom-right. The clamping when they overlap is the engine's, because it
    // depends on the box.
    simple!("roundRect", |a| {
        let n = a.len().saturating_sub(5);
        let r = |i: usize| f32_arg(a, 5 + i);
        let radii = match n {
            0 => [0.0; 4],
            1 => [r(0); 4],
            2 => [r(0), r(1), r(0), r(1)],
            3 => [r(0), r(1), r(2), r(1)],
            _ => [r(0), r(1), r(2), r(3)],
        };
        Op2D::RoundRect {
            x: f32_arg(a, 1),
            y: f32_arg(a, 2),
            w: f32_arg(a, 3),
            h: f32_arg(a, 4),
            radii,
        }
    });
    // `ellipse` with all eight arguments. The four-argument `ellipse` above
    // stays: SDL and .NET emit it.
    simple!("ellipseFull", |a| Op2D::EllipseFull {
        x: f32_arg(a, 1),
        y: f32_arg(a, 2),
        rx: f32_arg(a, 3),
        ry: f32_arg(a, 4),
        rotation: f32_arg(a, 5),
        start: f32_arg(a, 6),
        end: f32_arg(a, 7),
        ccw: bool_arg(a, 8),
    });
    simple!("fillWithRule", |a| Op2D::FillWithRule(str_arg(a, 1)));
    simple!("clipWithRule", |a| Op2D::ClipWithRule(str_arg(a, 1)));

    // `fill()`, `fill(fillRule)`, `fill(path)`, `fill(path, fillRule)` — one
    // IDL name with four forms, told apart by what argument 1 IS. A page writes
    // whichever it means and the overload resolves here, where the argument is
    // still an object; splitting them into four host functions would push that
    // choice onto every frontend instead.
    macro_rules! overloaded {
        ($name:literal, $plain:expr, $ruled:expr, $pathed:expr) => {
            vm.register_host_fn(
                "web:canvas",
                $name,
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let target = target_of(args.first());
                    let plain: fn() -> Op2D = $plain;
                    let ruled: fn(String) -> Op2D = $ruled;
                    let pathed: fn(PathDef, String) -> Op2D = $pathed;
                    let op = match path_arg(args.get(1)) {
                        // `fill(path)` / `fill(path, fillRule)`
                        Some(path) => {
                            let rule = if args.len() > 2 {
                                str_arg(args, 2)
                            } else {
                                "nonzero".to_string()
                            };
                            pathed(path, rule)
                        }
                        None if args.len() > 1 => ruled(str_arg(args, 1)),
                        None => plain(),
                    };
                    apply(&target, op);
                    Value::Null
                }),
            );
        };
    }
    overloaded!(
        "fill",
        || Op2D::Fill,
        Op2D::FillWithRule,
        Op2D::FillPath
    );
    overloaded!(
        "clip",
        || Op2D::Clip,
        Op2D::ClipWithRule,
        Op2D::ClipPath
    );
    // `stroke` has no fill rule — a stroke has no inside to decide about.
    vm.register_host_fn(
        "web:canvas",
        "stroke",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            match path_arg(args.get(1)) {
                Some(path) => apply(&target, Op2D::StrokePath(path)),
                None => apply(&target, Op2D::Stroke),
            }
            Value::Null
        }),
    );
    // `ctx.addPath(path)` folds a prebuilt shape into the context's own path.
    vm.register_host_fn(
        "web:canvas",
        "appendPath",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = target_of(args.first());
            if let Some(path) = path_arg(args.get(1)) {
                apply(&target, Op2D::AppendPath(path));
            }
            Value::Null
        }),
    );
    simple!("reset", |_a| Op2D::Reset);
    simple!("drawFocusIfNeeded", |a| Op2D::DrawFocusIfNeeded(bool_arg(
        a, 1
    )));

    // `fillText(text, x, y, maxWidth)` — a FOURTH argument changes the
    // operation: the string is condensed to fit rather than clipped.
    simple!("fillTextMaxWidth", |a| Op2D::FillTextMaxWidth(
        str_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ));
    simple!("strokeTextMaxWidth", |a| Op2D::StrokeTextMaxWidth(
        str_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ));

    // ── The half that ASKS ───────────────────────────────────────────────
    //
    // Every one of these was unreachable, because `apply` returns nothing and
    // there was no other way back across the seam. They go through `query`,
    // which answers `Absent` for a target that names no surface — so a caller
    // can tell "there is no canvas" from "the answer is zero", which is the
    // whole reason questions do not travel as ops.
    macro_rules! ask {
        ($name:literal, $build:expr, $render:expr) => {
            vm.register_host_fn(
                "web:canvas",
                $name,
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let target = target_of(args.first());
                    let build: fn(&[Value]) -> Query2D = $build;
                    let render: fn(Query2DValue) -> Value = $render;
                    render(query(&target, build(args)))
                }),
            );
        };
    }

    /// A `bool` answer, or `null` when there is no surface.
    fn as_bool(v: Query2DValue) -> Value {
        match v {
            Query2DValue::Bool(b) => Value::Bool(b),
            _ => Value::Null,
        }
    }
    fn as_text(v: Query2DValue) -> Value {
        match v {
            Query2DValue::Text(t) => Value::String(t.into()),
            _ => Value::Null,
        }
    }
    fn as_floats(v: Query2DValue) -> Value {
        match v {
            Query2DValue::Floats(f) => Value::Object(vybe_runtime::heap::alloc(
                Object::new_array(f.into_iter().map(|x| Value::F64(x as f64)).collect()),
            )),
            _ => Value::Null,
        }
    }

    // `isPointInPath(x, y [, rule])` or `isPointInPath(path, x, y [, rule])`.
    // The path form asks about the path GIVEN and leaves the context's current
    // path alone, which is the reason a page reaches for a `Path2D` at all.
    ask!(
        "isPointInPath",
        |a| match path_arg(a.get(1)) {
            Some(path) => Query2D::IsPointInPathOf {
                path,
                x: f32_arg(a, 2),
                y: f32_arg(a, 3),
                rule: if a.len() > 4 {
                    str_arg(a, 4)
                } else {
                    "nonzero".to_string()
                },
            },
            None => Query2D::IsPointInPath {
                x: f32_arg(a, 1),
                y: f32_arg(a, 2),
                rule: if a.len() > 3 {
                    str_arg(a, 3)
                } else {
                    "nonzero".to_string()
                },
            },
        },
        as_bool
    );
    ask!(
        "isPointInStroke",
        |a| match path_arg(a.get(1)) {
            Some(path) => Query2D::IsPointInStrokeOf {
                path,
                x: f32_arg(a, 2),
                y: f32_arg(a, 3),
            },
            None => Query2D::IsPointInStroke {
                x: f32_arg(a, 1),
                y: f32_arg(a, 2),
            },
        },
        as_bool
    );
    ask!("isContextLost", |_a| Query2D::IsContextLost, as_bool);
    ask!("getLineDash", |_a| Query2D::GetLineDash, as_floats);

    // `getTransform()` answers a `DOMMatrix`. Its six named members are what a
    // page reads; the object carries them under the IDL's own names.
    ask!("getTransform", |_a| Query2D::GetTransform, |v| {
        let Query2DValue::Matrix(m) = v else {
            return Value::Null;
        };
        let mut o = Object::new();
        o.properties
            .insert("__type".into(), Value::String("DOMMatrix".into()));
        for (name, value) in [
            ("a", m[0]),
            ("b", m[1]),
            ("c", m[2]),
            ("d", m[3]),
            ("e", m[4]),
            ("f", m[5]),
        ] {
            o.properties.insert(name.into(), Value::F64(value as f64));
        }
        Value::Object(vybe_runtime::heap::alloc(o))
    });

    // `getImageData(sx, sy, sw, sh)` — the same `ImageData` shape
    // `createImageData` answers, so a page can hand one straight back to
    // `putImageData`. STRAIGHT RGBA, not premultiplied: §4.12.5 is explicit,
    // and a premultiplied byte array read back would darken every translucent
    // pixel a little more on each round trip.
    ask!(
        "getImageData",
        |a| Query2D::GetImageData {
            sx: a.get(1).map(|v| v.as_i32()).unwrap_or(0),
            sy: a.get(2).map(|v| v.as_i32()).unwrap_or(0),
            sw: a.get(3).map(|v| v.as_i32().max(0) as u32).unwrap_or(0),
            sh: a.get(4).map(|v| v.as_i32().max(0) as u32).unwrap_or(0),
        },
        |v| {
            let Query2DValue::Pixels {
                data,
                width,
                height,
            } = v
            else {
                return Value::Null;
            };
            let bytes = Object::new_array(data.into_iter().map(|b| Value::I32(b as i32)).collect());
            let mut o = Object::new();
            o.properties
                .insert("__type".into(), Value::String("ImageData".into()));
            o.properties
                .insert("data".into(), Value::Object(vybe_runtime::heap::alloc(bytes)));
            o.properties.insert("width".into(), Value::I32(width as i32));
            o.properties
                .insert("height".into(), Value::I32(height as i32));
            Value::Object(vybe_runtime::heap::alloc(o))
        }
    );

    // `canvas.toDataURL(type, quality)` — the only route the spec gives a page
    // to its own pixels, because a page cannot be handed a file path.
    ask!(
        "toDataURL",
        |a| Query2D::ToDataUrl {
            mime: if a.len() > 1 {
                str_arg(a, 1)
            } else {
                "image/png".to_string()
            },
            quality: a.get(2).map(|v| v.as_f64() as f32),
        },
        as_text
    );
    // `canvas.toBlob` — the encoded bytes. Handed over as a byte array rather
    // than a `Blob`, which is a type this platform does not have yet.
    ask!(
        "toBlob",
        |a| Query2D::ToBlob {
            mime: if a.len() > 1 {
                str_arg(a, 1)
            } else {
                "image/png".to_string()
            },
            quality: a.get(2).map(|v| v.as_f64() as f32),
        },
        |v| match v {
            Query2DValue::Bytes(b) => Value::Object(vybe_runtime::heap::alloc(
                Object::new_array(b.into_iter().map(|x| Value::I32(x as i32)).collect()),
            )),
            _ => Value::Null,
        }
    );

    // `getContextAttributes()` — what the context was created with.
    ask!(
        "getContextAttributes",
        |_a| Query2D::GetContextAttributes,
        |v| {
            let Query2DValue::ContextAttributes {
                alpha,
                desynchronized,
                color_space,
                color_type,
                will_read_frequently,
            } = v
            else {
                return Value::Null;
            };
            let mut o = Object::new();
            o.properties.insert(
                "__type".into(),
                Value::String("CanvasRenderingContext2DSettings".into()),
            );
            o.properties.insert("alpha".into(), Value::Bool(alpha));
            o.properties
                .insert("desynchronized".into(), Value::Bool(desynchronized));
            o.properties
                .insert("colorSpace".into(), Value::String(color_space.into()));
            o.properties
                .insert("colorType".into(), Value::String(color_type.into()));
            o.properties.insert(
                "willReadFrequently".into(),
                Value::Bool(will_read_frequently),
            );
            Value::Object(vybe_runtime::heap::alloc(o))
        }
    );

    // The string attributes, read back. §4.12.5 requires every one of these to
    // serialize, and until now a page could set them and never ask.
    macro_rules! read_attr {
        ($name:literal, $which:expr) => {
            ask!(
                $name,
                |_a| Query2D::GetStringAttribute($which),
                as_text
            );
        };
    }
    read_attr!("getFont", StringAttribute::Font);
    read_attr!("getFillStyle", StringAttribute::FillStyle);
    read_attr!("getStrokeStyle", StringAttribute::StrokeStyle);
    read_attr!("getFilter", StringAttribute::Filter);
    read_attr!(
        "getGlobalCompositeOperation",
        StringAttribute::GlobalCompositeOperation
    );
    read_attr!(
        "getImageSmoothingQuality",
        StringAttribute::ImageSmoothingQuality
    );
    read_attr!("getShadowColor", StringAttribute::ShadowColor);
    read_attr!("getDirection", StringAttribute::Direction);
    read_attr!("getLetterSpacing", StringAttribute::LetterSpacing);
    read_attr!("getWordSpacing", StringAttribute::WordSpacing);
    read_attr!("getFontKerning", StringAttribute::FontKerning);
    read_attr!("getFontStretch", StringAttribute::FontStretch);
    read_attr!("getFontVariantCaps", StringAttribute::FontVariantCaps);
    read_attr!("getTextRendering", StringAttribute::TextRendering);
    read_attr!("getLang", StringAttribute::Lang);
    read_attr!("getTextAlign", StringAttribute::TextAlign);
    read_attr!("getTextBaseline", StringAttribute::TextBaseline);
    read_attr!("getLineCap", StringAttribute::LineCap);
    read_attr!("getLineJoin", StringAttribute::LineJoin);

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

    // ── Gradients and patterns ───────────────────────────────────────────
    //
    // These build an OBJECT the page holds, adds stops to, and later assigns to
    // `fillStyle`. That is why the seam has no handle for one: the definition
    // travels whole when it is assigned, so there is no registry to keep, no
    // lifetime to get wrong, and nothing to leak when a page drops a gradient
    // it never used.
    //
    // Not bound to a canvas, per the IDL — a gradient made on one context can
    // be used on another — so these take the context only to match the call
    // shape every other method here has.
    macro_rules! gradient_factory {
        ($name:literal, $kind:literal, $coords:expr) => {
            vm.register_host_fn(
                "web:canvas",
                $name,
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let read: fn(&[Value]) -> Vec<f32> = $coords;
                    let mut o = Object::new();
                    o.properties
                        .insert("__type".into(), Value::String("CanvasGradient".into()));
                    o.properties
                        .insert("__kind".into(), Value::String($kind.into()));
                    o.properties.insert(
                        "__coords".into(),
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                            read(args).into_iter().map(|v| Value::F64(v as f64)).collect(),
                        ))),
                    );
                    // The stops, in the order `addColorStop` adds them.
                    o.properties.insert(
                        "__stops".into(),
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new()))),
                    );
                    Value::Object(vybe_runtime::heap::alloc(o))
                }),
            );
        };
    }
    gradient_factory!("createLinearGradient", "linear", |a| vec![
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4)
    ]);
    gradient_factory!("createRadialGradient", "radial", |a| vec![
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3),
        f32_arg(a, 4),
        f32_arg(a, 5),
        f32_arg(a, 6)
    ]);
    // §4.12.5 puts the ANGLE first for a conic gradient, unlike the other two.
    gradient_factory!("createConicGradient", "conic", |a| vec![
        f32_arg(a, 1),
        f32_arg(a, 2),
        f32_arg(a, 3)
    ]);

    // `gradient.addColorStop(offset, color)` — mutates the gradient object.
    // The receiver here is the GRADIENT, not a context.
    vm.register_host_fn(
        "web:canvas",
        "addColorStop",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(g)) = args.first() else {
                return Value::Null;
            };
            let offset = args.get(1).map(|v| v.as_f64() as f32).unwrap_or(0.0);
            let color = str_arg(args, 2);
            let lock = g.lock().unwrap();
            if let Some(Value::Object(stops)) = lock.properties.get("__stops") {
                let mut stops = stops.lock().unwrap();
                if let ObjectKind::Array(ref mut list) = stops.kind {
                    let mut stop = Object::new();
                    stop.properties
                        .insert("offset".into(), Value::F64(offset as f64));
                    stop.properties
                        .insert("color".into(), Value::String(color.into()));
                    list.push(Value::Object(vybe_runtime::heap::alloc(stop)));
                }
            }
            Value::Null
        }),
    );

    // `createPattern(image, repetition)` — the image's pixels travel with it,
    // because a pattern is used long after the element that supplied them may
    // have changed.
    vm.register_host_fn(
        "web:canvas",
        "createPattern",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some((pixels, width, height)) = image_data_arg(args.get(1)) else {
                // §4.12.5: a source with no usable pixels answers null rather
                // than an unusable pattern.
                return Value::Null;
            };
            let mut o = Object::new();
            o.properties
                .insert("__type".into(), Value::String("CanvasPattern".into()));
            o.properties.insert(
                "__pixels".into(),
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                    pixels.into_iter().map(|b| Value::I32(b as i32)).collect(),
                ))),
            );
            o.properties.insert("__width".into(), Value::I32(width as i32));
            o.properties
                .insert("__height".into(), Value::I32(height as i32));
            o.properties.insert(
                "__repetition".into(),
                Value::String(
                    if args.len() > 2 {
                        str_arg(args, 2)
                    } else {
                        "repeat".to_string()
                    }
                    .into(),
                ),
            );
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // ── Path2D ───────────────────────────────────────────────────────────
    //
    // A path built once and used many times. It exists because the context's
    // own path is CONSUMED — `fill()` leaves it in place but `clip()` throws it
    // away, and `isPointInPath` cannot be asked about a shape the context is
    // only half-way through describing.
    //
    // The object carries its own segments, so nothing is registered on the
    // engine side: the operations travel when the path is USED. A page that
    // builds a path and drops it has cost the engine nothing.
    vm.register_host_fn(
        "web:canvas",
        "createPath2D",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let mut o = Object::new();
            o.properties
                .insert("__type".into(), Value::String("Path2D".into()));
            o.properties.insert(
                "__ops".into(),
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new()))),
            );
            let created = Value::Object(vybe_runtime::heap::alloc(o));
            // `new Path2D(otherPath)` — the copy constructor. Its segments are
            // COPIED, so later edits to either path leave the other alone.
            if let Some(source) = path_arg(args.first()) {
                for op in &source.ops {
                    copy_path_op(Some(&created), op);
                }
            }
            created
        }),
    );

    // The path-building methods, on a `Path2D` rather than on a context. Same
    // IDL names — a `Path2D` and a context build paths identically, which is
    // why the interface is worth having — so they are told apart by the
    // RECEIVER, which the vtable supplies as argument 0.
    macro_rules! path_op {
        ($name:literal, $op:literal, $fields:expr) => {
            vm.register_host_fn(
                "web:canvas",
                $name,
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let read: fn(&[Value]) -> Vec<(&'static str, Value)> = $fields;
                    push_path_op(args.first(), $op, &read(args));
                    Value::Null
                }),
            );
        };
    }
    let num = |a: &[Value], i: usize| Value::F64(f32_arg(a, i) as f64);
    let _ = num;
    path_op!("pathClosePath", "closePath", |_a| vec![]);
    path_op!("pathMoveTo", "moveTo", |a| vec![
        ("x", Value::F64(f32_arg(a, 1) as f64)),
        ("y", Value::F64(f32_arg(a, 2) as f64))
    ]);
    path_op!("pathLineTo", "lineTo", |a| vec![
        ("x", Value::F64(f32_arg(a, 1) as f64)),
        ("y", Value::F64(f32_arg(a, 2) as f64))
    ]);
    path_op!("pathQuadraticCurveTo", "quadraticCurveTo", |a| vec![
        ("cx", Value::F64(f32_arg(a, 1) as f64)),
        ("cy", Value::F64(f32_arg(a, 2) as f64)),
        ("x", Value::F64(f32_arg(a, 3) as f64)),
        ("y", Value::F64(f32_arg(a, 4) as f64))
    ]);
    path_op!("pathBezierCurveTo", "bezierCurveTo", |a| vec![
        ("cx1", Value::F64(f32_arg(a, 1) as f64)),
        ("cy1", Value::F64(f32_arg(a, 2) as f64)),
        ("cx2", Value::F64(f32_arg(a, 3) as f64)),
        ("cy2", Value::F64(f32_arg(a, 4) as f64)),
        ("x", Value::F64(f32_arg(a, 5) as f64)),
        ("y", Value::F64(f32_arg(a, 6) as f64))
    ]);
    path_op!("pathArcTo", "arcTo", |a| vec![
        ("x1", Value::F64(f32_arg(a, 1) as f64)),
        ("y1", Value::F64(f32_arg(a, 2) as f64)),
        ("x2", Value::F64(f32_arg(a, 3) as f64)),
        ("y2", Value::F64(f32_arg(a, 4) as f64)),
        ("radius", Value::F64(f32_arg(a, 5) as f64))
    ]);
    path_op!("pathRect", "rect", |a| vec![
        ("x", Value::F64(f32_arg(a, 1) as f64)),
        ("y", Value::F64(f32_arg(a, 2) as f64)),
        ("w", Value::F64(f32_arg(a, 3) as f64)),
        ("h", Value::F64(f32_arg(a, 4) as f64))
    ]);
    path_op!("pathArc", "arc", |a| vec![
        ("x", Value::F64(f32_arg(a, 1) as f64)),
        ("y", Value::F64(f32_arg(a, 2) as f64)),
        ("r", Value::F64(f32_arg(a, 3) as f64)),
        ("start", Value::F64(f32_arg(a, 4) as f64)),
        ("end", Value::F64(f32_arg(a, 5) as f64)),
        ("ccw", Value::Bool(bool_arg(a, 6)))
    ]);
    path_op!("pathEllipse", "ellipse", |a| vec![
        ("x", Value::F64(f32_arg(a, 1) as f64)),
        ("y", Value::F64(f32_arg(a, 2) as f64)),
        ("rx", Value::F64(f32_arg(a, 3) as f64)),
        ("ry", Value::F64(f32_arg(a, 4) as f64)),
        ("rotation", Value::F64(f32_arg(a, 5) as f64)),
        ("start", Value::F64(f32_arg(a, 6) as f64)),
        ("end", Value::F64(f32_arg(a, 7) as f64)),
        ("ccw", Value::Bool(bool_arg(a, 8)))
    ]);
    // The same short-form radii rule `roundRect` has on a context.
    path_op!("pathRoundRect", "roundRect", |a| {
        let n = a.len().saturating_sub(5);
        let r = |i: usize| f32_arg(a, 5 + i);
        let radii = match n {
            0 => [0.0; 4],
            1 => [r(0); 4],
            2 => [r(0), r(1), r(0), r(1)],
            3 => [r(0), r(1), r(2), r(1)],
            _ => [r(0), r(1), r(2), r(3)],
        };
        vec![
            ("x", Value::F64(f32_arg(a, 1) as f64)),
            ("y", Value::F64(f32_arg(a, 2) as f64)),
            ("w", Value::F64(f32_arg(a, 3) as f64)),
            ("h", Value::F64(f32_arg(a, 4) as f64)),
            ("r0", Value::F64(radii[0] as f64)),
            ("r1", Value::F64(radii[1] as f64)),
            ("r2", Value::F64(radii[2] as f64)),
            ("r3", Value::F64(radii[3] as f64)),
        ]
    });

    // `path.addPath(other)` — folds one path's segments into another.
    vm.register_host_fn(
        "web:canvas",
        "addPath",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(source) = path_arg(args.get(1)) else {
                return Value::Null;
            };
            for op in &source.ops {
                copy_path_op(args.first(), op);
            }
            Value::Null
        }),
    );

    // ── Wire the IDL attributes ──────────────────────────────────────────
    //
    // Last, because it reads back the indices of the functions registered
    // above. An attribute whose function is missing is left UNWIRED rather
    // than pointed at index zero: a wrong index would call some other host
    // function with a canvas as its first argument, which is a failure that
    // does something instead of nothing.
    let _ = ATTR_FNS.set(
        ATTRIBUTES
            .iter()
            .map(|(attr, getter, setter)| {
                let look = |name: &str| {
                    if name.is_empty() {
                        return None;
                    }
                    vm.host_registry
                        .get(&("web:canvas".to_string(), name.to_string()))
                        .copied()
                };
                (*attr, look(getter), look(setter))
            })
            .collect::<Vec<_>>(),
    );
}
