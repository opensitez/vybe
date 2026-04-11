//! Standalone canvas demo — proves the toolkit independence.
//!
//! Builds a `tiny_skia::Pixmap`, wraps it in a `TinySkiaCanvas`, draws a
//! few shapes via the generic `Canvas` trait, and saves the result as
//! `target/canvas_demo.png`.
//!
//! Run with:
//! ```sh
//! cargo run -p vybe_widgets --example canvas_demo
//! ```
//!
//! This binary depends ONLY on `vybe_widgets` and `tiny_skia` (which is
//! already a `vybe_widgets` dependency, re-exported as
//! `vybe_widgets::Pixmap`). No `vybe_host`, no `vybe_bytecode`, no .NET
//! wrapper layer.

use vybe_widgets::canvas::{Canvas, TinySkiaCanvas, RecordingCanvas, Color, LineCap, LineJoin};
use vybe_widgets::{Pixmap, Form, Canvas as CanvasWidget, FontSystem, SwashCache};
use vybe_widgets::layout::{LayoutRect, RenderContext, PanelWidget};

fn main() {
    let mut pixmap = Pixmap::new(800, 600).expect("create pixmap");
    pixmap.fill(tiny_skia::Color::WHITE);

    // ── Path 1: paint directly via TinySkiaCanvas ──────────────────────
    {
        let mut canvas = TinySkiaCanvas::new(&mut pixmap);

        // Filled red square in the top-left.
        canvas.set_fill_color(Color::rgb(220, 50, 50));
        canvas.fill_rect(20.0, 20.0, 120.0, 120.0);

        // Stroked outline around it.
        canvas.set_stroke_color(Color::rgb(40, 40, 40));
        canvas.set_line_width(3.0);
        canvas.stroke_rect(20.0, 20.0, 120.0, 120.0);

        // Path: a green diagonal line with rounded caps.
        canvas.set_stroke_color(Color::rgb(60, 180, 60));
        canvas.set_line_width(8.0);
        canvas.set_line_cap(LineCap::Round);
        canvas.set_line_join(LineJoin::Round);
        canvas.begin_path();
        canvas.move_to(180.0, 40.0);
        canvas.line_to(380.0, 240.0);
        canvas.stroke();

        // Bezier curve in blue.
        canvas.set_stroke_color(Color::rgb(50, 100, 220));
        canvas.set_line_width(4.0);
        canvas.begin_path();
        canvas.move_to(420.0, 40.0);
        canvas.bezier_curve_to(520.0, 40.0, 520.0, 240.0, 620.0, 240.0);
        canvas.stroke();

        // Filled ellipse.
        canvas.set_fill_color(Color::rgba(180, 60, 220, 180));
        canvas.begin_path();
        canvas.ellipse(700.0, 80.0, 60.0, 40.0);
        canvas.fill();
    }

    // ── Path 2: build a recording, then replay onto the same pixmap ────
    //
    // This is exactly how the host bridge works: user code (or VM
    // bytecode) appends to a RecordingCanvas; the form's render loop
    // replays it onto whatever live backend is active. Here we use the
    // same TinySkiaCanvas backend, just constructed against a fresh
    // borrow of the pixmap.
    {
        let mut recording = RecordingCanvas::new();
        recording.set_fill_color(Color::rgb(255, 200, 0));
        recording.fill_rect(20.0, 300.0, 760.0, 60.0);

        recording.set_stroke_color(Color::rgb(200, 100, 0));
        recording.set_line_width(2.0);
        recording.stroke_rect(20.0, 300.0, 760.0, 60.0);

        // Some text-shape sketch lines (text rendering is no-op at the
        // canvas-trait level — see TinySkiaCanvas::fill_text comment).
        recording.set_fill_color(Color::rgb(40, 40, 40));
        for i in 0..10 {
            let x = 40.0 + (i as f32) * 70.0;
            recording.fill_rect(x, 320.0, 40.0, 20.0);
        }

        let mut canvas = TinySkiaCanvas::new(&mut pixmap);
        recording.replay(&mut canvas);
    }

    // ── Save the trait-direct demo ──────────────────────────────────────
    let out_path = "target/canvas_demo.png";
    pixmap.save_png(out_path).expect("save png");
    println!("wrote {}", out_path);

    // ── Path 3: Canvas as a real PanelWidget on a Form ──────────────────
    //
    // Drop a `CanvasWidget` into a Form, paint into its underlying
    // RecordingCanvas via with_canvas, then render the form once and
    // save the result. This exercises the same code path the host
    // bridge will use when running VM forms — Form.render walks the
    // child widgets, each Canvas widget replays its recording onto the
    // active TinySkiaCanvas.
    let mut form_pixmap = Pixmap::new(800, 600).expect("create form pixmap");
    form_pixmap.fill(tiny_skia::Color::from_rgba8(245, 245, 250, 255));

    let mut form = Form::new("Canvas Widget Demo");
    form.set_rect(LayoutRect::new(0.0, 0.0, 800.0, 600.0));

    let mut canvas_widget = CanvasWidget::new()
        .with_name("art")
        .with_background(Color::WHITE);

    canvas_widget.with_canvas(|c| {
        // Draw a coordinate grid (widget-relative — origin is the
        // canvas widget's top-left, not the pixmap's).
        c.set_stroke_color(Color::rgb(220, 220, 220));
        c.set_line_width(1.0);
        for i in 0..=10 {
            let p = i as f32 * 60.0;
            c.begin_path();
            c.move_to(p, 0.0);
            c.line_to(p, 400.0);
            c.stroke();
            c.begin_path();
            c.move_to(0.0, p);
            c.line_to(600.0, p);
            c.stroke();
        }
        // Draw a triangle.
        c.set_fill_color(Color::rgba(50, 100, 220, 200));
        c.begin_path();
        c.move_to(300.0, 50.0);
        c.line_to(450.0, 250.0);
        c.line_to(150.0, 250.0);
        c.close_path();
        c.fill();

        // Stroke the same triangle.
        c.set_stroke_color(Color::rgb(20, 50, 120));
        c.set_line_width(3.0);
        c.set_line_join(LineJoin::Miter);
        c.begin_path();
        c.move_to(300.0, 50.0);
        c.line_to(450.0, 250.0);
        c.line_to(150.0, 250.0);
        c.close_path();
        c.stroke();
    });

    form.add_control(canvas_widget, 100.0, 100.0, 600.0, 400.0);

    // Render the form into the pixmap. cosmic-text needs a FontSystem
    // even though we don't use any text — it's part of RenderContext.
    let mut font_system = FontSystem::new();
    let mut swash_cache = SwashCache::new();
    let mut ctx = RenderContext {
        pixmap: &mut form_pixmap,
        font_system: &mut font_system,
        swash_cache: &mut swash_cache,
        scale: 1.0,
    };
    form.render(&mut ctx);

    let form_out = "target/canvas_widget_demo.png";
    form_pixmap.save_png(form_out).expect("save form png");
    println!("wrote {}", form_out);
}
