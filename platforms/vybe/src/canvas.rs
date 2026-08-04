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
//! 1. User code makes some framework-specific drawing call (e.g.
//!    `g.DrawLine(p, x1, y1, x2, y2)` in .NET).
//! 2. The matching framework wrapper compiles this into a sequence of
//!    `vybe:gui::canvas*` host calls (set stroke colour, set width,
//!    begin path, move to, line to, stroke). Each call passes a
//!    canvas-context handle as its first arg.
//! 3. The canvas-context handle is an Object stamped with
//!    `__control_name` (the source control's name, set by
//!    `getContext`). Framework wrappers may also stamp their own
//!    `__type` (`"Graphics"`, `"CanvasContext2D"`, etc.) for downcast
//!    purposes — the bridge doesn't care about the type tag.
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
//!
//! ## Naming
//!
//! The constructor is `getContext` (HTML5-canvas naming —
//! `element.getContext('2d')`). All paint/path/draw methods are
//! prefixed `canvas*`. The bridge is intentionally framework-neutral:
//! no `Pen`, `Brush`, `Graphics`, or other framework-specific concepts
//! leak in. Wrappers (`.NET System.Drawing.Graphics`, Flutter
//! `Canvas`, JS `getContext('2d')`) live OUTSIDE this module.

#[cfg(feature = "gui")]
mod canvas_impl {

    use crate::gui_state::GuiState;
    use std::sync::{Arc, Mutex};
    use vybe_runtime::value::Object;
    use vybe_runtime::{HostContext, VM, Value};
    use vybe_widgets::canvas::{Canvas, Color, Font, FontStyle, FontWeight, LineCap, LineJoin};

    pub fn register(vm: &mut VM, gui: Arc<Mutex<GuiState>>) {
        // ── Constructor: vybe:gui::getContext(controlName) ─────────────────
        //
        // Returns a small Object stamped `__type = "CanvasContext"` and
        // `__control_name = controlName`. Subsequent canvas host fns read
        // the name out of the handle to find the target canvas.
        //
        // Framework wrappers (the .NET `Control.CreateGraphics` body, the
        // future JS `Canvas.getContext('2d')`, the future Flutter
        // `RenderObject.canvas` bridge) call this and re-stamp `__type`
        // with their own framework-specific tag (`"Graphics"`,
        // `"CanvasRenderingContext2D"`, etc.) so user code can downcast.
        //
        // Side effect: ensures a `RecordingCanvas` exists for the control,
        // either as a Canvas widget's internal recording (if the widget is
        // already on the form) or as an entry in `overlay_canvases` (so
        // the form's render loop knows to replay it through `paint_overlay`).
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "getContext",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let ctrl_name = args.first().map(|v| format!("{}", v)).unwrap_or_default();
                    // Touch find_canvas_for_draw to ensure storage exists.
                    {
                        let _ = gui.lock().unwrap().find_canvas_mut(&ctrl_name);
                    }
                    let mut o = Object::new();
                    o.properties
                        .insert("__type".into(), Value::String(Arc::from("CanvasContext")));
                    o.properties.insert(
                        "__control_type".into(),
                        Value::String(Arc::from("CanvasContext")),
                    );
                    o.properties.insert(
                        "__control_name".into(),
                        Value::String(Arc::from(ctrl_name.to_lowercase().as_str())),
                    );
                    Value::Object(vybe_runtime::heap::alloc(o))
                }),
            );
        }

        // ── Paint state ────────────────────────────────────────────────────
        bind1_color(vm, &gui, "canvasSetFillColor", |c, col| {
            c.set_fill_color(col)
        });
        bind1_color(vm, &gui, "canvasSetStrokeColor", |c, col| {
            c.set_stroke_color(col)
        });
        bind1_f32(vm, &gui, "canvasSetLineWidth", |c, w| c.set_line_width(w));
        bind1_f32(vm, &gui, "canvasSetMiterLimit", |c, l| c.set_miter_limit(l));
        bind1_f32(vm, &gui, "canvasSetGlobalAlpha", |c, a| {
            c.set_global_alpha(a)
        });
        bind1_str(vm, &gui, "canvasSetLineCap", |c, s| {
            c.set_line_cap(parse_line_cap(s))
        });
        bind1_str(vm, &gui, "canvasSetLineJoin", |c, s| {
            c.set_line_join(parse_line_join(s))
        });

        // canvasSetFont(handle, family, size, bold, italic)
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasSetFont",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let name = handle_name(args.first());
                    let family = args
                        .get(1)
                        .map(|v| format!("{}", v))
                        .unwrap_or_else(|| "sans-serif".into());
                    let size = args.get(2).map(|v| v.as_f64() as f32).unwrap_or(12.0);
                    let bold = args
                        .get(3)
                        .map(|v| matches!(v, Value::Bool(true)) || v.as_f64() != 0.0)
                        .unwrap_or(false);
                    let italic = args
                        .get(4)
                        .map(|v| matches!(v, Value::Bool(true)) || v.as_f64() != 0.0)
                        .unwrap_or(false);
                    let font = Font {
                        family,
                        size,
                        weight: if bold {
                            FontWeight::Bold
                        } else {
                            FontWeight::Normal
                        },
                        style: if italic {
                            FontStyle::Italic
                        } else {
                            FontStyle::Normal
                        } };
                    gui.lock().unwrap().find_canvas_for_draw(&name).set_font(&font);
                    Value::Null
                }),
            );
        }

        // ── Path building ──────────────────────────────────────────────────
        bind0(vm, &gui, "canvasBeginPath", |c| c.begin_path());
        bind0(vm, &gui, "canvasClosePath", |c| c.close_path());
        bind2_f32(vm, &gui, "canvasMoveTo", |c, x, y| c.move_to(x, y));
        bind2_f32(vm, &gui, "canvasLineTo", |c, x, y| c.line_to(x, y));
        bind4_f32(vm, &gui, "canvasQuadTo", |c, cx, cy, x, y| {
            c.quadratic_curve_to(cx, cy, x, y)
        });
        bind6_f32(vm, &gui, "canvasBezierTo", |c, cx1, cy1, cx2, cy2, x, y| {
            c.bezier_curve_to(cx1, cy1, cx2, cy2, x, y)
        });
        bind4_f32(vm, &gui, "canvasRect", |c, x, y, w, h| c.rect(x, y, w, h));
        bind4_f32(vm, &gui, "canvasEllipse", |c, x, y, rx, ry| {
            c.ellipse(x, y, rx, ry)
        });

        // canvasArc(handle, x, y, r, start, end, ccw)
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasArc",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let name = handle_name(args.first());
                    let x = f32_arg(args, 1);
                    let y = f32_arg(args, 2);
                    let r = f32_arg(args, 3);
                    let start = f32_arg(args, 4);
                    let end = f32_arg(args, 5);
                    let ccw = args
                        .get(6)
                        .map(|v| matches!(v, Value::Bool(true)) || v.as_f64() != 0.0)
                        .unwrap_or(false);
                    gui.lock()
                        .unwrap()
                        .find_canvas_for_draw(&name)
                        .arc(x, y, r, start, end, ccw);
                    Value::Null
                }),
            );
        }

        // ── Drawing ────────────────────────────────────────────────────────
        bind0(vm, &gui, "canvasFill", |c| c.fill());
        bind0(vm, &gui, "canvasStroke", |c| c.stroke());
        bind4_f32(vm, &gui, "canvasFillRect", |c, x, y, w, h| {
            c.fill_rect(x, y, w, h)
        });
        bind4_f32(vm, &gui, "canvasStrokeRect", |c, x, y, w, h| {
            c.stroke_rect(x, y, w, h)
        });
        bind4_f32(vm, &gui, "canvasClearRect", |c, x, y, w, h| {
            c.clear_rect(x, y, w, h)
        });

        // ── Convenience composites ─────────────────────────────────────────
        //
        // These bundle a few canvas trait calls into one host fn. They're
        // useful for framework wrappers that take a bounding rect rather
        // than centre+radii (and would otherwise need arithmetic in the
        // body DSL). Each one is still expressible as a pure-trait
        // sequence — it's just packaged as a single host call so body
        // authors don't have to push-and-multiply.
        bind4_f32(vm, &gui, "canvasFillEllipseInRect", |c, x, y, w, h| {
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

        // canvasStrokeArcInRect(handle, x, y, w, h, startDeg, sweepDeg)
        // — bounding-rect arc with degree-based start/sweep. Wraps the
        // canvas trait's centre-based `arc` op, which uses radians and
        // start/end (not start/sweep). The conversion is: cx = x + w/2,
        // cy = y + h/2, r = avg(w, h)/2 (true elliptic arcs need
        // separate rx/ry — we approximate with a circular arc when
        // w == h, and a stretched circular arc otherwise).
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasStrokeArcInRect",
                Box::new(move |_ctx, args| {
                    let h = handle_name(args.first());
                    let x = f32_arg(args, 1);
                    let y = f32_arg(args, 2);
                    let w = f32_arg(args, 3);
                    let height = f32_arg(args, 4);
                    let start_deg = f32_arg(args, 5);
                    let sweep_deg = f32_arg(args, 6);
                    let cx = x + w * 0.5;
                    let cy = y + height * 0.5;
                    let r = (w + height) * 0.25;
                    let start = start_deg.to_radians();
                    let end = (start_deg + sweep_deg).to_radians();
                    let mut state = gui.lock().unwrap();
                    let canvas = state.find_canvas_for_draw(&h);
                    canvas.begin_path();
                    canvas.arc(cx, cy, r, start, end, false);
                    canvas.stroke();
                    Value::Null
                }),
            );
        }
        // canvasFillPieInRect(handle, x, y, w, h, startDeg, sweepDeg)
        // — same shape but builds a closed wedge (move to centre, arc,
        // close) and fills it.
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasFillPieInRect",
                Box::new(move |_ctx, args| {
                    let h = handle_name(args.first());
                    let x = f32_arg(args, 1);
                    let y = f32_arg(args, 2);
                    let w = f32_arg(args, 3);
                    let height = f32_arg(args, 4);
                    let start_deg = f32_arg(args, 5);
                    let sweep_deg = f32_arg(args, 6);
                    let cx = x + w * 0.5;
                    let cy = y + height * 0.5;
                    let r = (w + height) * 0.25;
                    let start = start_deg.to_radians();
                    let end = (start_deg + sweep_deg).to_radians();
                    let mut state = gui.lock().unwrap();
                    let canvas = state.find_canvas_for_draw(&h);
                    canvas.begin_path();
                    canvas.move_to(cx, cy);
                    canvas.line_to(cx + r * start.cos(), cy + r * start.sin());
                    canvas.arc(cx, cy, r, start, end, false);
                    canvas.close_path();
                    canvas.fill();
                    Value::Null
                }),
            );
        }
        // canvasStrokePieInRect — same as fill but stroke.
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasStrokePieInRect",
                Box::new(move |_ctx, args| {
                    let h = handle_name(args.first());
                    let x = f32_arg(args, 1);
                    let y = f32_arg(args, 2);
                    let w = f32_arg(args, 3);
                    let height = f32_arg(args, 4);
                    let start_deg = f32_arg(args, 5);
                    let sweep_deg = f32_arg(args, 6);
                    let cx = x + w * 0.5;
                    let cy = y + height * 0.5;
                    let r = (w + height) * 0.25;
                    let start = start_deg.to_radians();
                    let end = (start_deg + sweep_deg).to_radians();
                    let mut state = gui.lock().unwrap();
                    let canvas = state.find_canvas_for_draw(&h);
                    canvas.begin_path();
                    canvas.move_to(cx, cy);
                    canvas.line_to(cx + r * start.cos(), cy + r * start.sin());
                    canvas.arc(cx, cy, r, start, end, false);
                    canvas.close_path();
                    canvas.stroke();
                    Value::Null
                }),
            );
        }
        // canvasClearAll(handle, r, g, b, a) — fill the entire canvas
        // bounds with the given colour. We don't actually know the
        // canvas's bounding rect at this layer (the canvas is a free
        // surface; bounds belong to the widget), so the implementation
        // does a fill_rect over an arbitrarily large area starting at
        // (0,0). Tests / runners that care about exact bounds can issue
        // canvasClearRect(handle, x, y, w, h) directly.
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasClearAll",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let h = handle_name(args.first());
                    let r = (f32_arg(args, 1).clamp(0.0, 255.0)) as u8;
                    let g_ = (f32_arg(args, 2).clamp(0.0, 255.0)) as u8;
                    let b = (f32_arg(args, 3).clamp(0.0, 255.0)) as u8;
                    let a = (f32_arg(args, 4).clamp(0.0, 255.0)) as u8;
                    let mut state = gui.lock().unwrap();
                    let canvas = state.find_canvas_for_draw(&h);
                    canvas.set_fill_color(Color::rgba(r, g_, b, a));
                    canvas.fill_rect(0.0, 0.0, 100_000.0, 100_000.0);
                    Value::Null
                }),
            );
        }

        // canvasFillText(handle, text, x, y)
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasFillText",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let name = handle_name(args.first());
                    let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                    let x = f32_arg(args, 2);
                    let y = f32_arg(args, 3);
                    gui.lock()
                        .unwrap()
                        .find_canvas_for_draw(&name)
                        .fill_text(&text, x, y);
                    Value::Null
                }),
            );
        }
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasStrokeText",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let name = handle_name(args.first());
                    let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                    let x = f32_arg(args, 2);
                    let y = f32_arg(args, 3);
                    gui.lock()
                        .unwrap()
                        .find_canvas_for_draw(&name)
                        .stroke_text(&text, x, y);
                    Value::Null
                }),
            );
        }
        // canvasDrawImage(handle, pixels, srcW, srcH, x, y, dstW, dstH)
        //
        // `pixels` is a dense array of RGBA bytes, srcW*srcH*4 of them — the
        // shape a software renderer produces. This is the blit path: a guest
        // computes a whole frame into a byte buffer and hands it over once,
        // instead of issuing a host call per primitive. Doom's renderer is
        // exactly this, and so is any SDL program using a surface rather than
        // the 2D renderer API.
        //
        // Layer 1 was already complete — `Image`, `Canvas::draw_image`,
        // `RecordingCanvas` DrawImage recording/replay and the tiny_skia
        // `draw_pixmap` blit all existed; only this decode was missing.
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasDrawImage",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let name = handle_name(args.first());
                    let src_w = args.get(2).map(|v| v.as_f64() as u32).unwrap_or(0);
                    let src_h = args.get(3).map(|v| v.as_f64() as u32).unwrap_or(0);
                    if src_w == 0 || src_h == 0 {
                        return Value::Null;
                    }
                    let x = f32_arg(args, 4);
                    let y = f32_arg(args, 5);
                    // Destination size defaults to the source size, so a
                    // 1:1 blit needs only six arguments.
                    let dst_w = args
                        .get(6)
                        .map(|v| v.as_f64() as f32)
                        .filter(|v| *v > 0.0)
                        .unwrap_or(src_w as f32);
                    let dst_h = args
                        .get(7)
                        .map(|v| v.as_f64() as f32)
                        .filter(|v| *v > 0.0)
                        .unwrap_or(src_h as f32);

                    let needed = (src_w as usize) * (src_h as usize) * 4;
                    let mut bytes: Vec<u8> = Vec::with_capacity(needed);
                    if let Some(Value::Object(obj)) = args.get(1) {
                        let o = obj.lock().unwrap();
                        if let vybe_runtime::value::ObjectKind::Array(items) = &o.kind {
                            for v in items.iter().take(needed) {
                                bytes.push(v.as_f64() as u8);
                            }
                        }
                    }
                    if bytes.len() < needed {
                        // Short buffer: pad rather than panic. `Image::from_rgba`
                        // debug-asserts the exact length, and a guest that
                        // miscounts should get a visibly wrong frame, not a
                        // host abort.
                        bytes.resize(needed, 0);
                    }

                    let img = vybe_widgets::canvas::Image::from_rgba(src_w, src_h, bytes);
                    gui.lock()
                        .unwrap()
                        .find_canvas_for_draw(&name)
                        .draw_image(&img, x, y, dst_w, dst_h);
                    Value::Null
                }),
            );
        }

        // ── Dashed strokes ─────────────────────────────────────────────────
        //
        // Framework wrappers may want to pass variable-length dash
        // arrays. Rather than introduce an array-encoding convention
        // through the host bridge, we expose a small set of fixed-arity
        // setters: 0 (solid), 2 (simple dash), 4 (dash-dot), 6 (dash-dot-dot).
        // Wrappers like .NET's `DashStyle` enum map their value to one
        // of these. Body sequences calling these don't need to know the
        // length encoding.
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasSetLineDashSolid",
                Box::new(move |_ctx, args| {
                    let h = handle_name(args.first());
                    gui.lock().unwrap().find_canvas_for_draw(&h).set_line_dash(&[]);
                    Value::Null
                }),
            );
        }
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasSetLineDash2",
                Box::new(move |_ctx, args| {
                    let h = handle_name(args.first());
                    let d0 = f32_arg(args, 1);
                    let d1 = f32_arg(args, 2);
                    gui.lock()
                        .unwrap()
                        .find_canvas_for_draw(&h)
                        .set_line_dash(&[d0, d1]);
                    Value::Null
                }),
            );
        }
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasSetLineDash4",
                Box::new(move |_ctx, args| {
                    let h = handle_name(args.first());
                    let d0 = f32_arg(args, 1);
                    let d1 = f32_arg(args, 2);
                    let d2 = f32_arg(args, 3);
                    let d3 = f32_arg(args, 4);
                    gui.lock()
                        .unwrap()
                        .find_canvas_for_draw(&h)
                        .set_line_dash(&[d0, d1, d2, d3]);
                    Value::Null
                }),
            );
        }
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasSetLineDash6",
                Box::new(move |_ctx, args| {
                    let h = handle_name(args.first());
                    let mut intervals = [0.0f32; 6];
                    for i in 0..6 {
                        intervals[i] = f32_arg(args, 1 + i);
                    }
                    gui.lock()
                        .unwrap()
                        .find_canvas_for_draw(&h)
                        .set_line_dash(&intervals);
                    Value::Null
                }),
            );
        }
        bind1_f32(vm, &gui, "canvasSetLineDashOffset", |c, o| {
            c.set_line_dash_offset(o)
        });

        // canvasApplyPenDashStyle(handle, enumValue)
        //
        // Convenience for framework wrappers that have a `DashStyle` enum.
        // Maps the .NET `System.Drawing.Drawing2D.DashStyle` integer values
        // to a fixed dash pattern, and applies it to the canvas state. The
        // .NET DashStyle enum is used as-is by other framework wrappers
        // (Flutter, JS canvas API on the web side, …) — they pass the same
        // integer values.
        //
        // Mapping (matches .NET):
        //   0 = Solid       → []
        //   1 = Dash        → [6, 4]
        //   2 = Dot         → [2, 4]
        //   3 = DashDot     → [6, 4, 2, 4]
        //   4 = DashDotDot  → [6, 4, 2, 4, 2, 4]
        //   5 = Custom      → no-op (pattern is already on the canvas via
        //                     SetLineDashN)
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "canvasApplyPenDashStyle",
                Box::new(move |_ctx, args| {
                    let h = handle_name(args.first());
                    let style = f32_arg(args, 1) as i32;
                    let pattern: &[f32] = match style {
                        0 => &[],
                        1 => &[6.0, 4.0],
                        2 => &[2.0, 4.0],
                        3 => &[6.0, 4.0, 2.0, 4.0],
                        4 => &[6.0, 4.0, 2.0, 4.0, 2.0, 4.0],
                        _ => return Value::Null };
                    gui.lock()
                        .unwrap()
                        .find_canvas_for_draw(&h)
                        .set_line_dash(pattern);
                    Value::Null
                }),
            );
        }

        // ── Clipping ───────────────────────────────────────────────────────
        bind0(vm, &gui, "canvasClip", |c| c.clip());
        bind0(vm, &gui, "canvasResetClip", |c| c.reset_clip());

        // ── State stack ────────────────────────────────────────────────────
        bind0(vm, &gui, "canvasSave", |c| c.save());
        bind0(vm, &gui, "canvasRestore", |c| c.restore());

        // ── Transforms ─────────────────────────────────────────────────────
        bind2_f32(vm, &gui, "canvasTranslate", |c, x, y| c.translate(x, y));
        bind1_f32(vm, &gui, "canvasRotate", |c, rad| c.rotate(rad));
        // Convenience: rotate by degrees instead of radians. Saves the
        // dotnet body sequences from doing the *π/180 conversion.
        bind1_f32(vm, &gui, "canvasRotateDegrees", |c, deg| {
            c.rotate(deg.to_radians())
        });
        bind2_f32(vm, &gui, "canvasScale", |c, sx, sy| c.scale(sx, sy));
        bind6_f32(
            vm,
            &gui,
            "canvasTransform",
            |c, m11, m12, m21, m22, dx, dy| c.transform(m11, m12, m21, m22, dx, dy),
        );
        bind0(vm, &gui, "canvasResetTransform", |c| c.reset_transform());

        // ── new_Canvas constructor (parallel to new_Button etc.) ───────────
        //
        // Constructs a Canvas widget on the form (or in the host's stub
        // table). The dotnet `PaintBox` / `PictureBox` class wrapper
        // declares this as its `widget_host_fn`.
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "new_Canvas",
                Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
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
                    o.properties
                        .insert("__control_type".into(), Value::String(Arc::from("Canvas")));
                    o.properties.insert(
                        "__control_name".into(),
                        Value::String(Arc::from(name.as_str())),
                    );
                    o.properties
                        .insert("__type".into(), Value::String(Arc::from("Canvas")));
                    o.properties
                        .insert("name".into(), Value::String(Arc::from(name.as_str())));
                    Value::Object(vybe_runtime::heap::alloc(o))
                }),
            );
        }
        
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "sdlFillRect",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let name = handle_name(args.first());
                    let mut use_full = false;
                    let mut x = 0.0;
                    let mut y = 0.0;
                    let mut w = 0.0;
                    let mut h = 0.0;

                    if let Some(rect) = args.get(1) {
                        if matches!(rect, Value::Null) || matches!(rect, Value::I32(0)) || matches!(rect, Value::F64(0.0)) {
                            use_full = true;
                        } else if let Value::Object(obj) = rect {
                            // `&rect` is a POINTER CELL — `{ __ref_kind: "cell",
                            // __value: <struct> }` (see `primitives/pointers.rs`).
                            // Unwrap it before looking for fields, or every
                            // `SDL_FillRect(s, &r, c)` decoded to w=h=0 and
                            // filled nothing while still "succeeding".
                            let obj = match obj.lock().unwrap().properties.get("__value") {
                                Some(Value::Object(inner)) => inner.clone(),
                                _ => obj.clone() };
                            let obj_lock = obj.lock().unwrap();
                            if let Some(Value::Object(base_obj)) = obj_lock.properties.get("__base") {
                                let base_lock = base_obj.lock().unwrap();
                                if let vybe_runtime::value::ObjectKind::Array(ref arr) = base_lock.kind {
                                    x = arr.get(0).map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                    y = arr.get(1).map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                    w = arr.get(2).map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                    h = arr.get(3).map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                } else {
                                    x = base_lock.properties.get("x").map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                    y = base_lock.properties.get("y").map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                    w = base_lock.properties.get("w").map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                    h = base_lock.properties.get("h").map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                }
                            } else {
                                x = obj_lock.properties.get("x").map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                y = obj_lock.properties.get("y").map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                w = obj_lock.properties.get("w").map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                                h = obj_lock.properties.get("h").map(|v| v.as_f64()).unwrap_or(0.0) as f32;
                            }
                        }
                    } else {
                        use_full = true;
                    }

                    let mut state = gui.lock().unwrap();
                    if use_full {
                        let w_str = state.get_property(&name, "width");
                        w = w_str.parse().unwrap_or(0.0);
                        let h_str = state.get_property(&name, "height");
                        h = h_str.parse().unwrap_or(0.0);
                    }

                    if let Some(c) = args.get(2) {
                        // Read through f64, NOT `as_i32`: a packed colour with
                        // alpha (0xFFRRGGBB = 4283879648 for cyan) exceeds
                        // i32::MAX, and `as i32` SATURATES to 0x7FFFFFFF — which
                        // unpacks to r=g=b=0xFF, so every colour rendered white.
                        let c_val = c.as_f64() as u32;
                        let r = ((c_val >> 16) & 0xFF) as u8;
                        let g = ((c_val >> 8) & 0xFF) as u8;
                        let b = (c_val & 0xFF) as u8;
                        let a = ((c_val >> 24) & 0xFF) as u8;
                        let color = vybe_widgets::canvas::Color::rgba(r, g, b, if a == 0 { 255 } else { a });
                        state.find_canvas_for_draw(&name).set_fill_color(color);
                    }

                    state.find_canvas_for_draw(&name).fill_rect(x, y, w, h);
                    Value::Null
                }),
            );
        }
        
        // sdlPresent(surface) — SDL's frame boundary.
        //
        // SDL is immediate-mode and never clears its surface; the program just
        // redraws. A RecordingCanvas is retained-mode, so without a boundary an
        // animated program's commands grow without bound and every frame paints
        // over the last. This marks the canvas so the NEXT draw starts a fresh
        // frame, which keeps the presented frame on screen until it is replaced.
        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "sdlPresent",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let name = handle_name(args.first());
                    let mut state = gui.lock().unwrap();
                    let resolved = state.resolve_control_name(&name);
                    state.pending_clear.insert(resolved);
                    Value::Null
                }),
            );
        }

        {
            let gui = gui.clone();
            vm.register_host_fn(
                "vybe:gui",
                "sdlDrawLine",
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let name = handle_name(args.first());
                    if vybe_widgets::canvas::trace_enabled() {
                        eprintln!("[sdlDrawLine] argc={} arg0={:?} name={name:?}", args.len(), args.first());
                    }
                    let x1 = f32_arg(args, 1);
                    let y1 = f32_arg(args, 2);
                    let x2 = f32_arg(args, 3);
                    let y2 = f32_arg(args, 4);

                    let mut state = gui.lock().unwrap();
                    if let Some(c) = args.get(5) {
                        // Read through f64, NOT `as_i32`: a packed colour with
                        // alpha (0xFFRRGGBB = 4283879648 for cyan) exceeds
                        // i32::MAX, and `as i32` SATURATES to 0x7FFFFFFF — which
                        // unpacks to r=g=b=0xFF, so every colour rendered white.
                        let c_val = c.as_f64() as u32;
                        let r = ((c_val >> 16) & 0xFF) as u8;
                        let g = ((c_val >> 8) & 0xFF) as u8;
                        let b = (c_val & 0xFF) as u8;
                        let a = ((c_val >> 24) & 0xFF) as u8;
                        let color = vybe_widgets::canvas::Color::rgba(r, g, b, if a == 0 { 255 } else { a });
                        state.find_canvas_for_draw(&name).set_stroke_color(color);
                        state.find_canvas_for_draw(&name).set_line_width(1.0);
                    }
                    
                    let canvas = state.find_canvas_for_draw(&name);
                    canvas.begin_path();
                    canvas.move_to(x1, y1);
                    canvas.line_to(x2, y2);
                    canvas.stroke();
                    Value::Null
                }),
            );
        }
    }

    // ─── Argument helpers ──────────────────────────────────────────────────────

    fn f32_arg(args: &[Value], idx: usize) -> f32 {
        args.get(idx).map(|v| v.as_f64() as f32).unwrap_or(0.0)
    }

    /// Control name for a canvas target, accepting either shape a caller can
    /// hold: a CanvasContext handle from `getContext`, or the control NAME
    /// itself.
    ///
    /// The name form is what `getContext` already accepts (it stringifies its
    /// argument), and it is what SDL passes — `SDL_CreateWindow` stores the
    /// `<window>_surface` Canvas control's NAME as `sdl_surface`. Rejecting it
    /// here returned an empty string, so `find_canvas_for_draw("")` recorded every
    /// `sdlDrawLine` into a canvas belonging to no control, which nothing
    /// paints: the call succeeded and the line never appeared.
    fn handle_name(arg: Option<&Value>) -> String {
        if vybe_widgets::canvas::trace_enabled() {
            let shape = match arg {
                None => "None".to_string(),
                Some(Value::Object(o)) => {
                    let g = o.lock().unwrap();
                    format!("Object(props={:?})", g.properties.keys().collect::<Vec<_>>())
                }
                Some(v) => format!("{:?}", v) };
            eprintln!("[handle_name] arg={shape}");
        }
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

    fn parse_line_cap(s: &str) -> LineCap {
        match s.to_ascii_lowercase().as_str() {
            "round" => LineCap::Round,
            "square" => LineCap::Square,
            _ => LineCap::Butt }
    }

    fn parse_line_join(s: &str) -> LineJoin {
        match s.to_ascii_lowercase().as_str() {
            "round" => LineJoin::Round,
            "bevel" => LineJoin::Bevel,
            _ => LineJoin::Miter }
    }

    // ─── Bind helpers ──────────────────────────────────────────────────────────
    //
    // Most canvas host fns follow one of a handful of stereotyped shapes:
    // "0 args after handle" (Fill, Stroke, BeginPath, …), "1 f32" (LineWidth,
    // Rotate, …), "2 f32" (MoveTo, Translate, …), etc. These helpers cut the
    // boilerplate. Each one builds the closure, locks GuiState, looks up the
    // canvas, and forwards.

    fn bind0<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
    where
        F: Fn(&mut dyn Canvas) + Send + Sync + 'static,
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            name,
            Box::new(move |_ctx, args| {
                let h = handle_name(args.first());
                f(gui.lock().unwrap().find_canvas_for_draw(&h));
                Value::Null
            }),
        );
    }

    fn bind1_f32<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
    where
        F: Fn(&mut dyn Canvas, f32) + Send + Sync + 'static,
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            name,
            Box::new(move |_ctx, args| {
                let h = handle_name(args.first());
                f(gui.lock().unwrap().find_canvas_for_draw(&h), f32_arg(args, 1));
                Value::Null
            }),
        );
    }

    fn bind2_f32<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
    where
        F: Fn(&mut dyn Canvas, f32, f32) + Send + Sync + 'static,
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            name,
            Box::new(move |_ctx, args| {
                let h = handle_name(args.first());
                f(
                    gui.lock().unwrap().find_canvas_for_draw(&h),
                    f32_arg(args, 1),
                    f32_arg(args, 2),
                );
                Value::Null
            }),
        );
    }

    fn bind4_f32<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
    where
        F: Fn(&mut dyn Canvas, f32, f32, f32, f32) + Send + Sync + 'static,
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            name,
            Box::new(move |_ctx, args| {
                let h = handle_name(args.first());
                f(
                    gui.lock().unwrap().find_canvas_for_draw(&h),
                    f32_arg(args, 1),
                    f32_arg(args, 2),
                    f32_arg(args, 3),
                    f32_arg(args, 4),
                );
                Value::Null
            }),
        );
    }

    fn bind6_f32<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
    where
        F: Fn(&mut dyn Canvas, f32, f32, f32, f32, f32, f32) + Send + Sync + 'static,
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            name,
            Box::new(move |_ctx, args| {
                let h = handle_name(args.first());
                f(
                    gui.lock().unwrap().find_canvas_for_draw(&h),
                    f32_arg(args, 1),
                    f32_arg(args, 2),
                    f32_arg(args, 3),
                    f32_arg(args, 4),
                    f32_arg(args, 5),
                    f32_arg(args, 6),
                );
                Value::Null
            }),
        );
    }

    fn bind1_color<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
    where
        F: Fn(&mut dyn Canvas, Color) + Send + Sync + 'static,
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            name,
            Box::new(move |_ctx, args| {
                let h = handle_name(args.first());
                let r = (f32_arg(args, 1).clamp(0.0, 255.0)) as u8;
                let g = (f32_arg(args, 2).clamp(0.0, 255.0)) as u8;
                let b = (f32_arg(args, 3).clamp(0.0, 255.0)) as u8;
                // Alpha defaults to 255 if omitted (lets dotnet wrappers pass
                // RGB without forcing them to specify the alpha channel).
                let a = if args.len() > 4 {
                    (f32_arg(args, 4).clamp(0.0, 255.0)) as u8
                } else {
                    255
                };
                f(
                    gui.lock().unwrap().find_canvas_for_draw(&h),
                    Color::rgba(r, g, b, a),
                );
                Value::Null
            }),
        );
    }

    fn bind1_str<F>(vm: &mut VM, gui: &Arc<Mutex<GuiState>>, name: &str, f: F)
    where
        F: Fn(&mut dyn Canvas, &str) + Send + Sync + 'static,
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            name,
            Box::new(move |_ctx, args| {
                let h = handle_name(args.first());
                let s = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                f(gui.lock().unwrap().find_canvas_for_draw(&h), &s);
                Value::Null
            }),
        );
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
pub fn register(_vm: &mut vybe_runtime::VM) {}
