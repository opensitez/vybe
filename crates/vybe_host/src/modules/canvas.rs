//! `vybe:gui::canvas*` — host bridge between the VM and the
//! `vybe_widgets::canvas::Canvas` trait.
//!
//! This is the **transport layer** for canvas drawing. Every host fn
//! here is a thin forwarder: lock `GuiState`, look up the target
//! `RecordingCanvas` via [`GuiState::find_canvas_mut`], call the
//! corresponding [`Canvas`] trait method. Zero drawing logic lives
//! here — the trait IS the interface, recording captures the calls,
//! and the form's render loop replays them.
//!
//! ## How a canvas call lands
//!
//! 1. User code does `g.DrawLine(p, x1, y1, x2, y2)`.
//! 2. The .NET wrapper layer compiles this into a sequence of
//!    `vybe:gui::canvas*` host calls (set stroke colour, set width,
//!    begin path, move to, line to, stroke). Each call passes the
//!    Graphics handle as its first arg.
//! 3. The Graphics handle is an Object stamped with `__control_name`
//!    (the source control's name, set by `createGraphics`).
//! 4. Each canvas host fn here reads the name out of the handle, asks
//!    `GuiState::find_canvas_mut` for the target canvas, and forwards.
//!
//! ## What the canvas IS
//!
//! For a Canvas WIDGET on the form: the widget's own internal
//! `RecordingCanvas` (via `Canvas::canvas_mut`).
//!
//! For any other control (Form, Button, Label, …): the per-control
//! overlay recording in `GuiState::overlay_canvases`. The form's
//! render loop replays each overlay through the matching widget's
//! `paint_overlay` hook each frame.
//!
//! Both flow through the same trait — only the storage location
//! differs.

#[cfg(feature = "gui")]
mod canvas_impl {

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;
use vybe_widgets::canvas::{Canvas, Color, LineCap, LineJoin, Font, FontWeight, FontStyle};
use crate::gui_state::GuiState;

pub fn register(vm: &mut VM, gui: Arc<Mutex<GuiState>>) {
    // ── Constructor: vybe:gui::createGraphics(controlName) ─────────────
    //
    // Returns a small Object stamped `__type = "Graphics"` and
    // `__control_name = controlName`. Subsequent canvas host fns read
    // the name out of the handle to find the target canvas.
    //
    // Side effect: ensures a `RecordingCanvas` exists for the control,
    // either as a Canvas widget's internal recording (if the widget is
    // already on the form) or as an entry in `overlay_canvases` (so
    // the form's render loop knows to replay it through `paint_overlay`).
    {
        let gui = gui.clone();
        vm.register_host_fn("vybe:gui", "createGraphics", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let ctrl_name = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            // Touch find_canvas_mut to ensure storage exists.
            { let _ = gui.lock().unwrap().find_canvas_mut(&ctrl_name); }
            let mut o = Object::new();
            o.properties.insert("__type".into(), Value::String(Arc::from("Graphics")));
            o.properties.insert("__control_type".into(), Value::String(Arc::from("Graphics")));
            o.properties.insert("__control_name".into(), Value::String(Arc::from(ctrl_name.to_lowercase().as_str())));
            Value::Object(Arc::new(Mutex::new(o)))
        }));
    }

    // ── Paint state ────────────────────────────────────────────────────
    bind1_color(vm, &gui, "canvasSetFillColor",   |c, col| c.set_fill_color(col));
    bind1_color(vm, &gui, "canvasSetStrokeColor", |c, col| c.set_stroke_color(col));
    bind1_f32(vm, &gui, "canvasSetLineWidth",   |c, w| c.set_line_width(w));
    bind1_f32(vm, &gui, "canvasSetMiterLimit",  |c, l| c.set_miter_limit(l));
    bind1_f32(vm, &gui, "canvasSetGlobalAlpha", |c, a| c.set_global_alpha(a));
    bind1_str(vm, &gui, "canvasSetLineCap",     |c, s| c.set_line_cap(parse_line_cap(s)));
    bind1_str(vm, &gui, "canvasSetLineJoin",    |c, s| c.set_line_join(parse_line_join(s)));

    // canvasSetFont(handle, family, size, bold, italic)
    {
        let gui = gui.clone();
        vm.register_host_fn("vybe:gui", "canvasSetFont", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let name = handle_name(args.first());
            let family = args.get(1).map(|v| format!("{}", v)).unwrap_or_else(|| "sans-serif".into());
            let size = args.get(2).map(|v| v.as_f64() as f32).unwrap_or(12.0);
            let bold = args.get(3).map(|v| matches!(v, Value::Bool(true)) || v.as_f64() != 0.0).unwrap_or(false);
            let italic = args.get(4).map(|v| matches!(v, Value::Bool(true)) || v.as_f64() != 0.0).unwrap_or(false);
            let font = Font {
                family,
                size,
                weight: if bold { FontWeight::Bold } else { FontWeight::Normal },
                style: if italic { FontStyle::Italic } else { FontStyle::Normal },
            };
            gui.lock().unwrap().find_canvas_mut(&name).set_font(&font);
            Value::Null
        }));
    }

    // ── Path building ──────────────────────────────────────────────────
    bind0(vm, &gui, "canvasBeginPath", |c| c.begin_path());
    bind0(vm, &gui, "canvasClosePath", |c| c.close_path());
    bind2_f32(vm, &gui, "canvasMoveTo",  |c, x, y| c.move_to(x, y));
    bind2_f32(vm, &gui, "canvasLineTo",  |c, x, y| c.line_to(x, y));
    bind4_f32(vm, &gui, "canvasQuadTo",  |c, cx, cy, x, y| c.quadratic_curve_to(cx, cy, x, y));
    bind6_f32(vm, &gui, "canvasBezierTo", |c, cx1, cy1, cx2, cy2, x, y| c.bezier_curve_to(cx1, cy1, cx2, cy2, x, y));
    bind4_f32(vm, &gui, "canvasRect",    |c, x, y, w, h| c.rect(x, y, w, h));
    bind4_f32(vm, &gui, "canvasEllipse", |c, x, y, rx, ry| c.ellipse(x, y, rx, ry));

    // canvasArc(handle, x, y, r, start, end, ccw)
    {
        let gui = gui.clone();
        vm.register_host_fn("vybe:gui", "canvasArc", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let name = handle_name(args.first());
            let x = f32_arg(args, 1);
            let y = f32_arg(args, 2);
            let r = f32_arg(args, 3);
            let start = f32_arg(args, 4);
            let end = f32_arg(args, 5);
            let ccw = args.get(6).map(|v| matches!(v, Value::Bool(true)) || v.as_f64() != 0.0).unwrap_or(false);
            gui.lock().unwrap().find_canvas_mut(&name).arc(x, y, r, start, end, ccw);
            Value::Null
        }));
    }

    // ── Drawing ────────────────────────────────────────────────────────
    bind0(vm, &gui, "canvasFill", |c| c.fill());
    bind0(vm, &gui, "canvasStroke", |c| c.stroke());
    bind4_f32(vm, &gui, "canvasFillRect",   |c, x, y, w, h| c.fill_rect(x, y, w, h));
    bind4_f32(vm, &gui, "canvasStrokeRect", |c, x, y, w, h| c.stroke_rect(x, y, w, h));
    bind4_f32(vm, &gui, "canvasClearRect",  |c, x, y, w, h| c.clear_rect(x, y, w, h));

    // ── Convenience composites ─────────────────────────────────────────
    //
    // These bundle a few canvas trait calls into one host fn. They're
    // useful for framework wrappers (`.NET Graphics.DrawEllipse` takes a
    // bounding rect, not centre+radii) that would otherwise need
    // arithmetic in the body DSL. Each one is still expressible as a
    // pure-trait sequence — it's just packaged as a single host call so
    // body authors don't have to push-and-multiply.
    bind4_f32(vm, &gui, "canvasFillEllipseInRect",   |c, x, y, w, h| {
        let cx = x + w * 0.5;
        let cy = y + h * 0.5;
        c.begin_path();
        c.ellipse(cx, cy, w * 0.5, h * 0.5);
        c.fill();
    });
    bind4_f32(vm, &gui, "canvasStrokeEllipseInRect", |c, x, y, w, h| {
        let cx = x + w * 0.5;
        let cy = y + h * 0.5;
        c.begin_path();
        c.ellipse(cx, cy, w * 0.5, h * 0.5);
        c.stroke();
    });
    // canvasClearAll(handle, r, g, b, a) — fill the entire canvas
    // bounds with the given colour. We don't actually know the
    // canvas's bounding rect at this layer (the canvas is a free
    // surface; bounds belong to the widget), so the implementation
    // does a fill_rect over an arbitrarily large area starting at
    // (0,0). Tests / runners that care about exact bounds can issue
    // canvasClearRect(handle, x, y, w, h) directly.
    {
        let gui = gui.clone();
        vm.register_host_fn("vybe:gui", "canvasClearAll", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let h = handle_name(args.first());
            let r = (f32_arg(args, 1).clamp(0.0, 255.0)) as u8;
            let g_ = (f32_arg(args, 2).clamp(0.0, 255.0)) as u8;
            let b = (f32_arg(args, 3).clamp(0.0, 255.0)) as u8;
            let a = (f32_arg(args, 4).clamp(0.0, 255.0)) as u8;
            let mut state = gui.lock().unwrap();
            let canvas = state.find_canvas_mut(&h);
            canvas.set_fill_color(Color::rgba(r, g_, b, a));
            canvas.fill_rect(0.0, 0.0, 100_000.0, 100_000.0);
            Value::Null
        }));
    }

    // canvasFillText(handle, text, x, y)
    {
        let gui = gui.clone();
        vm.register_host_fn("vybe:gui", "canvasFillText", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let name = handle_name(args.first());
            let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let x = f32_arg(args, 2);
            let y = f32_arg(args, 3);
            gui.lock().unwrap().find_canvas_mut(&name).fill_text(&text, x, y);
            Value::Null
        }));
    }
    {
        let gui = gui.clone();
        vm.register_host_fn("vybe:gui", "canvasStrokeText", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let name = handle_name(args.first());
            let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let x = f32_arg(args, 2);
            let y = f32_arg(args, 3);
            gui.lock().unwrap().find_canvas_mut(&name).stroke_text(&text, x, y);
            Value::Null
        }));
    }
    // canvasDrawImage is left as a no-op until image loading is wired
    // through the host (Layer 1 has the trait method but the host
    // bridge doesn't yet decode an Image from a Value).
    {
        let gui = gui.clone();
        vm.register_host_fn("vybe:gui", "canvasDrawImage", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            // TODO: when a host-side Image type lands, decode args[1]
            // into a `vybe_widgets::canvas::Image` and call
            // `canvas.draw_image(&img, x, y, w, h)`.
            let _ = gui;
            Value::Null
        }));
    }

    // ── State stack ────────────────────────────────────────────────────
    bind0(vm, &gui, "canvasSave", |c| c.save());
    bind0(vm, &gui, "canvasRestore", |c| c.restore());

    // ── Transforms ─────────────────────────────────────────────────────
    bind2_f32(vm, &gui, "canvasTranslate", |c, x, y| c.translate(x, y));
    bind1_f32(vm, &gui, "canvasRotate",    |c, rad| c.rotate(rad));
    bind2_f32(vm, &gui, "canvasScale",     |c, sx, sy| c.scale(sx, sy));
    bind6_f32(vm, &gui, "canvasTransform", |c, m11, m12, m21, m22, dx, dy|
        c.transform(m11, m12, m21, m22, dx, dy));
    bind0(vm, &gui, "canvasResetTransform", |c| c.reset_transform());

    // ── new_Canvas constructor (parallel to new_Button etc.) ───────────
    //
    // Constructs a Canvas widget on the form (or in the host's stub
    // table). The dotnet `PaintBox` / `PictureBox` class wrapper
    // declares this as its `widget_host_fn`.
    {
        let gui = gui.clone();
        vm.register_host_fn("vybe:gui", "new_Canvas", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            use std::sync::atomic::{AtomicU32, Ordering};
            static COUNTER: AtomicU32 = AtomicU32::new(1);
            let id = COUNTER.fetch_add(1, Ordering::Relaxed);
            let name = format!("canvas_{}", id);
            // Create the actual Canvas widget on the form so the user's
            // recording lives in the right place. The widget is added
            // at (0,0) with default size; user code is expected to set
            // Left/Top/Width/Height via the dotnet wrapper's setters
            // (which mirror through controlSetProperty).
            {
                let mut g = gui.lock().unwrap();
                g.add_widget("Canvas", &name, "", 0, 0, 100, 100);
            }
            // Return a control-shaped Object for the dotnet wrapper to
            // copy `__control_name` etc. off.
            let mut o = Object::new();
            o.properties.insert("__control_type".into(), Value::String(Arc::from("Canvas")));
            o.properties.insert("__control_name".into(), Value::String(Arc::from(name.as_str())));
            o.properties.insert("__type".into(), Value::String(Arc::from("Canvas")));
            o.properties.insert("name".into(), Value::String(Arc::from(name.as_str())));
            Value::Object(Arc::new(Mutex::new(o)))
        }));
    }
}

// ─── Argument helpers ──────────────────────────────────────────────────────

fn f32_arg(args: &[Value], idx: usize) -> f32 {
    args.get(idx).map(|v| v.as_f64() as f32).unwrap_or(0.0)
}

fn handle_name(arg: Option<&Value>) -> String {
    if let Some(Value::Object(obj)) = arg {
        let o = obj.lock().unwrap();
        if let Some(v) = o.properties.get("__control_name") {
            return format!("{}", v).to_lowercase();
        }
    }
    String::new()
}

fn parse_line_cap(s: &str) -> LineCap {
    match s.to_ascii_lowercase().as_str() {
        "round" => LineCap::Round,
        "square" => LineCap::Square,
        _ => LineCap::Butt,
    }
}

fn parse_line_join(s: &str) -> LineJoin {
    match s.to_ascii_lowercase().as_str() {
        "round" => LineJoin::Round,
        "bevel" => LineJoin::Bevel,
        _ => LineJoin::Miter,
    }
}

// ─── Bind helpers ──────────────────────────────────────────────────────────
//
// Most canvas host fns follow one of a handful of stereotyped shapes:
// "0 args after handle" (Fill, Stroke, BeginPath, …), "1 f32" (LineWidth,
// Rotate, …), "2 f32" (MoveTo, Translate, …), etc. These helpers cut the
// boilerplate. Each one builds the closure, locks GuiState, looks up the
// canvas, and forwards.

fn bind0<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
where F: Fn(&mut dyn Canvas) + Send + Sync + 'static {
    let gui = gui.clone();
    vm.register_host_fn("vybe:gui", name, Box::new(move |_ctx, args| {
        let h = handle_name(args.first());
        f(gui.lock().unwrap().find_canvas_mut(&h));
        Value::Null
    }));
}

fn bind1_f32<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
where F: Fn(&mut dyn Canvas, f32) + Send + Sync + 'static {
    let gui = gui.clone();
    vm.register_host_fn("vybe:gui", name, Box::new(move |_ctx, args| {
        let h = handle_name(args.first());
        f(gui.lock().unwrap().find_canvas_mut(&h), f32_arg(args, 1));
        Value::Null
    }));
}

fn bind2_f32<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
where F: Fn(&mut dyn Canvas, f32, f32) + Send + Sync + 'static {
    let gui = gui.clone();
    vm.register_host_fn("vybe:gui", name, Box::new(move |_ctx, args| {
        let h = handle_name(args.first());
        f(gui.lock().unwrap().find_canvas_mut(&h), f32_arg(args, 1), f32_arg(args, 2));
        Value::Null
    }));
}

fn bind4_f32<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
where F: Fn(&mut dyn Canvas, f32, f32, f32, f32) + Send + Sync + 'static {
    let gui = gui.clone();
    vm.register_host_fn("vybe:gui", name, Box::new(move |_ctx, args| {
        let h = handle_name(args.first());
        f(
            gui.lock().unwrap().find_canvas_mut(&h),
            f32_arg(args, 1), f32_arg(args, 2), f32_arg(args, 3), f32_arg(args, 4),
        );
        Value::Null
    }));
}

fn bind6_f32<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
where F: Fn(&mut dyn Canvas, f32, f32, f32, f32, f32, f32) + Send + Sync + 'static {
    let gui = gui.clone();
    vm.register_host_fn("vybe:gui", name, Box::new(move |_ctx, args| {
        let h = handle_name(args.first());
        f(
            gui.lock().unwrap().find_canvas_mut(&h),
            f32_arg(args, 1), f32_arg(args, 2),
            f32_arg(args, 3), f32_arg(args, 4),
            f32_arg(args, 5), f32_arg(args, 6),
        );
        Value::Null
    }));
}

fn bind1_color<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
where F: Fn(&mut dyn Canvas, Color) + Send + Sync + 'static {
    let gui = gui.clone();
    vm.register_host_fn("vybe:gui", name, Box::new(move |_ctx, args| {
        let h = handle_name(args.first());
        let r = (f32_arg(args, 1).clamp(0.0, 255.0)) as u8;
        let g = (f32_arg(args, 2).clamp(0.0, 255.0)) as u8;
        let b = (f32_arg(args, 3).clamp(0.0, 255.0)) as u8;
        // Alpha defaults to 255 if omitted (lets dotnet wrappers pass
        // RGB without forcing them to specify the alpha channel).
        let a = if args.len() > 4 {
            (f32_arg(args, 4).clamp(0.0, 255.0)) as u8
        } else { 255 };
        f(gui.lock().unwrap().find_canvas_mut(&h), Color::rgba(r, g, b, a));
        Value::Null
    }));
}

fn bind1_str<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
where F: Fn(&mut dyn Canvas, &str) + Send + Sync + 'static {
    let gui = gui.clone();
    vm.register_host_fn("vybe:gui", name, Box::new(move |_ctx, args| {
        let h = handle_name(args.first());
        let s = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        f(gui.lock().unwrap().find_canvas_mut(&h), &s);
        Value::Null
    }));
}

} // mod canvas_impl

// ─── Public re-export with feature gating ──────────────────────────────────

#[cfg(feature = "gui")]
pub use canvas_impl::register;

/// Stub register fn for the non-`gui` build. Canvas host fns require
/// `vybe_widgets::canvas`, which only ships in the `gui`-featured build,
/// so non-GUI consumers get a no-op (and the old test fallback path
/// supplies its own stubs).
#[cfg(not(feature = "gui"))]
pub fn register(_vm: &mut vybe_bytecode::VM) {}
