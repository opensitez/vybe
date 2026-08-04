//! `RecordingCanvas` — a `Canvas` impl that captures every call as data.
//!
//! Two reasons recording is the persistent state of a canvas:
//!
//! 1. **Tests.** Asserting on `recording.commands` is exact-shape and
//!    deterministic — no flaky pixel comparisons.
//!
//! 2. **Persistent draws.** Most retained-mode drawing models (HTML5
//!    canvas, immediate-mode UIs that need redraw on resize, framework
//!    `Graphics` objects, …) want "drawings persist between paint
//!    cycles". Recording is the natural shape for this — the form's
//!    render loop replays the recording onto a fresh `TinySkiaCanvas`
//!    each frame.
//!
//! `replay(&mut other)` walks the captured commands and dispatches each
//! one to another `Canvas` impl. The dispatch is mechanical — one
//! variant per trait method, one match arm per variant.

use super::{Canvas, Color, Font, Image, LineCap, LineJoin};

/// One captured drawing primitive.
///
/// Variants mirror the [`Canvas`] trait method-for-method (one variant
/// per call). Adding a method to the trait is matched by adding one
/// variant here and one match arm in [`RecordingCanvas::replay`].
#[derive(Clone, Debug, PartialEq)]
pub enum DrawCmd {
    // Paint state
    SetFillColor(Color),
    SetStrokeColor(Color),
    SetLineWidth(f32),
    SetLineCap(LineCap),
    SetLineJoin(LineJoin),
    SetMiterLimit(f32),
    SetGlobalAlpha(f32),
    SetFont(Font),
    SetLineDash(Vec<f32>),
    SetLineDashOffset(f32),

    // Path building
    BeginPath,
    ClosePath,
    MoveTo(f32, f32),
    LineTo(f32, f32),
    QuadraticCurveTo {
        cx: f32,
        cy: f32,
        x: f32,
        y: f32,
    },
    BezierCurveTo {
        cx1: f32,
        cy1: f32,
        cx2: f32,
        cy2: f32,
        x: f32,
        y: f32,
    },
    Arc {
        x: f32,
        y: f32,
        r: f32,
        start: f32,
        end: f32,
        ccw: bool,
    },
    Rect {
        x: f32,
        y: f32,
        w: f32,
        h: f32,
    },
    Ellipse {
        x: f32,
        y: f32,
        rx: f32,
        ry: f32,
    },

    // Drawing
    Fill,
    Stroke,
    FillRect {
        x: f32,
        y: f32,
        w: f32,
        h: f32,
    },
    StrokeRect {
        x: f32,
        y: f32,
        w: f32,
        h: f32,
    },
    ClearRect {
        x: f32,
        y: f32,
        w: f32,
        h: f32,
    },
    FillText {
        text: String,
        x: f32,
        y: f32,
    },
    StrokeText {
        text: String,
        x: f32,
        y: f32,
    },
    DrawImage {
        image: Image,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
    },
    Clip,
    ResetClip,

    // State stack
    Save,
    Restore,

    // Transforms
    Translate(f32, f32),
    Rotate(f32),
    Scale(f32, f32),
    Transform {
        m11: f32,
        m12: f32,
        m21: f32,
        m22: f32,
        dx: f32,
        dy: f32,
    },
    ResetTransform,
}

/// A `Canvas` impl that captures every call as a [`DrawCmd`].
///
/// `RecordingCanvas` is what every drawing surface holds as its
/// authoritative state. The form's render loop calls
/// [`RecordingCanvas::replay`] each frame to paint the captured commands
/// onto whatever live backend is active.
#[derive(Clone, Debug, Default)]
pub struct RecordingCanvas {
    pub commands: Vec<DrawCmd>,
}

impl RecordingCanvas {
    /// Construct an empty recording.
    pub fn new() -> Self {
        Self::default()
    }

    /// Replay every captured command onto another canvas.
    ///
    /// The dispatch is mechanical — one match arm per [`DrawCmd`]
    /// variant calling the corresponding [`Canvas`] trait method on
    /// `target`. After replay, the recording is unchanged (callers can
    /// `clear()` if they want one-shot replay).
    pub fn replay(&self, target: &mut dyn Canvas) {
        for cmd in &self.commands {
            match cmd {
                // Paint state
                DrawCmd::SetFillColor(c) => target.set_fill_color(*c),
                DrawCmd::SetStrokeColor(c) => target.set_stroke_color(*c),
                DrawCmd::SetLineWidth(w) => target.set_line_width(*w),
                DrawCmd::SetLineCap(cap) => target.set_line_cap(*cap),
                DrawCmd::SetLineJoin(j) => target.set_line_join(*j),
                DrawCmd::SetMiterLimit(l) => target.set_miter_limit(*l),
                DrawCmd::SetGlobalAlpha(a) => target.set_global_alpha(*a),
                DrawCmd::SetFont(f) => target.set_font(f),
                DrawCmd::SetLineDash(intervals) => target.set_line_dash(intervals),
                DrawCmd::SetLineDashOffset(o) => target.set_line_dash_offset(*o),

                // Path building
                DrawCmd::BeginPath => target.begin_path(),
                DrawCmd::ClosePath => target.close_path(),
                DrawCmd::MoveTo(x, y) => target.move_to(*x, *y),
                DrawCmd::LineTo(x, y) => target.line_to(*x, *y),
                DrawCmd::QuadraticCurveTo { cx, cy, x, y } => {
                    target.quadratic_curve_to(*cx, *cy, *x, *y)
                }
                DrawCmd::BezierCurveTo {
                    cx1,
                    cy1,
                    cx2,
                    cy2,
                    x,
                    y,
                } => target.bezier_curve_to(*cx1, *cy1, *cx2, *cy2, *x, *y),
                DrawCmd::Arc {
                    x,
                    y,
                    r,
                    start,
                    end,
                    ccw,
                } => target.arc(*x, *y, *r, *start, *end, *ccw),
                DrawCmd::Rect { x, y, w, h } => target.rect(*x, *y, *w, *h),
                DrawCmd::Ellipse { x, y, rx, ry } => target.ellipse(*x, *y, *rx, *ry),

                // Drawing
                DrawCmd::Fill => target.fill(),
                DrawCmd::Stroke => target.stroke(),
                DrawCmd::FillRect { x, y, w, h } => target.fill_rect(*x, *y, *w, *h),
                DrawCmd::StrokeRect { x, y, w, h } => target.stroke_rect(*x, *y, *w, *h),
                DrawCmd::ClearRect { x, y, w, h } => target.clear_rect(*x, *y, *w, *h),
                DrawCmd::FillText { text, x, y } => target.fill_text(text, *x, *y),
                DrawCmd::StrokeText { text, x, y } => target.stroke_text(text, *x, *y),
                DrawCmd::DrawImage { image, x, y, w, h } => {
                    target.draw_image(image, *x, *y, *w, *h)
                }
                DrawCmd::Clip => target.clip(),
                DrawCmd::ResetClip => target.reset_clip(),

                // State stack
                DrawCmd::Save => target.save(),
                DrawCmd::Restore => target.restore(),

                // Transforms
                DrawCmd::Translate(x, y) => target.translate(*x, *y),
                DrawCmd::Rotate(rad) => target.rotate(*rad),
                DrawCmd::Scale(sx, sy) => target.scale(*sx, *sy),
                DrawCmd::Transform {
                    m11,
                    m12,
                    m21,
                    m22,
                    dx,
                    dy,
                } => target.transform(*m11, *m12, *m21, *m22, *dx, *dy),
                DrawCmd::ResetTransform => target.reset_transform(),
            }
        }
    }

    /// Drop every captured command. Used by callers that want one-shot
    /// replay (replay → clear → start fresh).
    /// Debug-only view of the recorded commands.
    pub fn commands_for_debug(&self) -> &[DrawCmd] {
        &self.commands
    }

    pub fn clear(&mut self) {
        self.commands.clear();
    }

    /// Number of captured commands.
    pub fn len(&self) -> usize {
        self.commands.len()
    }

    /// True if no commands have been captured yet.
    pub fn is_empty(&self) -> bool {
        self.commands.is_empty()
    }
}

impl Canvas for RecordingCanvas {
    // Paint state
    fn set_fill_color(&mut self, color: Color) {
        self.commands.push(DrawCmd::SetFillColor(color));
    }
    fn set_stroke_color(&mut self, color: Color) {
        self.commands.push(DrawCmd::SetStrokeColor(color));
    }
    fn set_line_width(&mut self, width: f32) {
        self.commands.push(DrawCmd::SetLineWidth(width));
    }
    fn set_line_cap(&mut self, cap: LineCap) {
        self.commands.push(DrawCmd::SetLineCap(cap));
    }
    fn set_line_join(&mut self, join: LineJoin) {
        self.commands.push(DrawCmd::SetLineJoin(join));
    }
    fn set_miter_limit(&mut self, limit: f32) {
        self.commands.push(DrawCmd::SetMiterLimit(limit));
    }
    fn set_global_alpha(&mut self, alpha: f32) {
        self.commands.push(DrawCmd::SetGlobalAlpha(alpha));
    }
    fn set_font(&mut self, font: &Font) {
        self.commands.push(DrawCmd::SetFont(font.clone()));
    }
    fn set_line_dash(&mut self, intervals: &[f32]) {
        self.commands.push(DrawCmd::SetLineDash(intervals.to_vec()));
    }
    fn set_line_dash_offset(&mut self, offset: f32) {
        self.commands.push(DrawCmd::SetLineDashOffset(offset));
    }

    // Path building
    fn begin_path(&mut self) {
        self.commands.push(DrawCmd::BeginPath);
    }
    fn close_path(&mut self) {
        self.commands.push(DrawCmd::ClosePath);
    }
    fn move_to(&mut self, x: f32, y: f32) {
        self.commands.push(DrawCmd::MoveTo(x, y));
    }
    fn line_to(&mut self, x: f32, y: f32) {
        self.commands.push(DrawCmd::LineTo(x, y));
    }
    fn quadratic_curve_to(&mut self, cx: f32, cy: f32, x: f32, y: f32) {
        self.commands
            .push(DrawCmd::QuadraticCurveTo { cx, cy, x, y });
    }
    fn bezier_curve_to(&mut self, cx1: f32, cy1: f32, cx2: f32, cy2: f32, x: f32, y: f32) {
        self.commands.push(DrawCmd::BezierCurveTo {
            cx1,
            cy1,
            cx2,
            cy2,
            x,
            y,
        });
    }
    fn arc(&mut self, x: f32, y: f32, r: f32, start: f32, end: f32, ccw: bool) {
        self.commands.push(DrawCmd::Arc {
            x,
            y,
            r,
            start,
            end,
            ccw,
        });
    }
    fn rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.commands.push(DrawCmd::Rect { x, y, w, h });
    }
    fn ellipse(&mut self, x: f32, y: f32, rx: f32, ry: f32) {
        self.commands.push(DrawCmd::Ellipse { x, y, rx, ry });
    }

    // Drawing
    fn fill(&mut self) {
        self.commands.push(DrawCmd::Fill);
    }
    fn stroke(&mut self) {
        self.commands.push(DrawCmd::Stroke);
    }
    fn fill_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.commands.push(DrawCmd::FillRect { x, y, w, h });
    }
    fn stroke_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.commands.push(DrawCmd::StrokeRect { x, y, w, h });
    }
    fn clear_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.commands.push(DrawCmd::ClearRect { x, y, w, h });
    }
    fn fill_text(&mut self, text: &str, x: f32, y: f32) {
        self.commands.push(DrawCmd::FillText {
            text: text.to_string(),
            x,
            y,
        });
    }
    fn stroke_text(&mut self, text: &str, x: f32, y: f32) {
        self.commands.push(DrawCmd::StrokeText {
            text: text.to_string(),
            x,
            y,
        });
    }
    fn draw_image(&mut self, img: &Image, x: f32, y: f32, w: f32, h: f32) {
        self.commands.push(DrawCmd::DrawImage {
            image: img.clone(),
            x,
            y,
            w,
            h,
        });
    }
    fn clip(&mut self) {
        self.commands.push(DrawCmd::Clip);
    }
    fn reset_clip(&mut self) {
        self.commands.push(DrawCmd::ResetClip);
    }

    // State stack
    fn save(&mut self) {
        self.commands.push(DrawCmd::Save);
    }
    fn restore(&mut self) {
        self.commands.push(DrawCmd::Restore);
    }

    // Transforms
    fn translate(&mut self, x: f32, y: f32) {
        self.commands.push(DrawCmd::Translate(x, y));
    }
    fn rotate(&mut self, rad: f32) {
        self.commands.push(DrawCmd::Rotate(rad));
    }
    fn scale(&mut self, sx: f32, sy: f32) {
        self.commands.push(DrawCmd::Scale(sx, sy));
    }
    fn transform(&mut self, m11: f32, m12: f32, m21: f32, m22: f32, dx: f32, dy: f32) {
        self.commands.push(DrawCmd::Transform {
            m11,
            m12,
            m21,
            m22,
            dx,
            dy,
        });
    }
    fn reset_transform(&mut self) {
        self.commands.push(DrawCmd::ResetTransform);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn captures_calls_in_order() {
        let mut c = RecordingCanvas::new();
        c.set_fill_color(Color::rgb(255, 0, 0));
        c.fill_rect(0.0, 0.0, 10.0, 10.0);
        c.set_stroke_color(Color::rgb(0, 0, 0));
        c.set_line_width(2.0);
        c.begin_path();
        c.move_to(0.0, 0.0);
        c.line_to(100.0, 100.0);
        c.stroke();

        assert_eq!(c.len(), 8);
        assert_eq!(c.commands[0], DrawCmd::SetFillColor(Color::rgb(255, 0, 0)));
        assert_eq!(
            c.commands[1],
            DrawCmd::FillRect {
                x: 0.0,
                y: 0.0,
                w: 10.0,
                h: 10.0
            }
        );
        assert_eq!(c.commands[7], DrawCmd::Stroke);
    }

    #[test]
    fn replay_dispatches_to_target() {
        let mut src = RecordingCanvas::new();
        src.set_fill_color(Color::rgb(1, 2, 3));
        src.fill_rect(4.0, 5.0, 6.0, 7.0);
        src.move_to(8.0, 9.0);

        let mut dst = RecordingCanvas::new();
        src.replay(&mut dst);

        assert_eq!(src.commands, dst.commands);
    }

    #[test]
    fn clear_drops_all_commands() {
        let mut c = RecordingCanvas::new();
        c.fill_rect(0.0, 0.0, 1.0, 1.0);
        assert_eq!(c.len(), 1);
        c.clear();
        assert!(c.is_empty());
    }
}
