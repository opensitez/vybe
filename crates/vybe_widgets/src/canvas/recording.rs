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
    SetImageSmoothing(bool),
    SetTextAlign(super::TextAlign),
    SetTextBaseline(super::TextBaseline),
    SetFont(Font),
    SetLineDash(Vec<f32>),
    SetLineDashOffset(f32),
    /// `fillStyle` / `strokeStyle` when they hold a gradient or a pattern.
    ///
    /// Recorded SEPARATELY from `SetFillColor` rather than replacing it: a
    /// colour is by far the common case, and collapsing the two would make
    /// every existing recording assertion carry a `Paint::Color(..)` wrapper
    /// for no gain. `set_fill_color` still records `SetFillColor`.
    SetFillPaint(super::Paint),
    SetStrokePaint(super::Paint),
    SetShadow(super::Shadow),

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
    /// The spec's full `ellipse()` — rotation and a start/end angle pair.
    ///
    /// Its own variant rather than a widened `Ellipse`, because the default
    /// `ellipse_arc` composes save/rotate/scale/arc/restore: recorded through
    /// that path an elliptical wedge would arrive as five commands and the
    /// recording would no longer say what was asked for. A recording is a
    /// transcript, so the call is the unit.
    EllipseArc {
        x: f32,
        y: f32,
        rx: f32,
        ry: f32,
        rotation: f32,
        start: f32,
        end: f32,
        ccw: bool,
    },

    // Drawing
    Fill,
    FillWithRule(super::FillRule),
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
    /// `putImageData` — a RAW write. Recorded as its own command rather than
    /// a `DrawImage` with `w`/`h` equal to the image's, because the two differ
    /// in what they honour, not in their arguments: this one ignores the
    /// transform, the clip, `globalAlpha` and the blend mode. Replayed as a
    /// `DrawImage` it would pick all four up from whatever state the replay
    /// had reached.
    PutImageData {
        image: Image,
        dx: f32,
        dy: f32,
    },
    DrawImage {
        image: Image,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
    },
    /// The nine-argument `drawImage` — the SOURCE rectangle is the point, and
    /// cropping at record time would throw away what the call said.
    DrawImageRect {
        image: Image,
        sx: f32,
        sy: f32,
        sw: f32,
        sh: f32,
        dx: f32,
        dy: f32,
        dw: f32,
        dh: f32,
    },
    Clip,
    ClipWithRule(super::FillRule),
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
                DrawCmd::SetImageSmoothing(on) => target.set_image_smoothing(*on),
                DrawCmd::SetTextAlign(a) => target.set_text_align(*a),
                DrawCmd::SetTextBaseline(b) => target.set_text_baseline(*b),
                DrawCmd::SetFont(f) => target.set_font(f),
                DrawCmd::SetLineDash(intervals) => target.set_line_dash(intervals),
                DrawCmd::SetLineDashOffset(o) => target.set_line_dash_offset(*o),
                DrawCmd::SetFillPaint(p) => target.set_fill_paint(p),
                DrawCmd::SetStrokePaint(p) => target.set_stroke_paint(p),
                DrawCmd::SetShadow(s) => target.set_shadow(s),

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
                DrawCmd::EllipseArc {
                    x,
                    y,
                    rx,
                    ry,
                    rotation,
                    start,
                    end,
                    ccw,
                } => target.ellipse_arc(*x, *y, *rx, *ry, *rotation, *start, *end, *ccw),

                // Drawing
                DrawCmd::Fill => target.fill(),
                DrawCmd::FillWithRule(rule) => target.fill_with_rule(*rule),
                DrawCmd::Stroke => target.stroke(),
                DrawCmd::FillRect { x, y, w, h } => target.fill_rect(*x, *y, *w, *h),
                DrawCmd::StrokeRect { x, y, w, h } => target.stroke_rect(*x, *y, *w, *h),
                DrawCmd::ClearRect { x, y, w, h } => target.clear_rect(*x, *y, *w, *h),
                DrawCmd::FillText { text, x, y } => target.fill_text(text, *x, *y),
                DrawCmd::StrokeText { text, x, y } => target.stroke_text(text, *x, *y),
                DrawCmd::PutImageData { image, dx, dy } => {
                    target.put_image_data(image, *dx, *dy)
                }
                DrawCmd::DrawImage { image, x, y, w, h } => {
                    target.draw_image(image, *x, *y, *w, *h)
                }
                DrawCmd::DrawImageRect {
                    image,
                    sx,
                    sy,
                    sw,
                    sh,
                    dx,
                    dy,
                    dw,
                    dh,
                } => target.draw_image_rect(image, *sx, *sy, *sw, *sh, *dx, *dy, *dw, *dh),
                DrawCmd::Clip => target.clip(),
                DrawCmd::ClipWithRule(rule) => target.clip_with_rule(*rule),
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

    /// Where the path being built currently ends.
    ///
    /// **Derived from the commands, not tracked beside them.** A `last_point`
    /// field would have to be updated by every path method and would be wrong
    /// the moment one forgot — the same two-copies problem the rest of this
    /// engine has been shedding. Scanning backwards cannot drift, and `arcTo`
    /// is the only caller.
    /// The font in effect — what `measureText` and `fillText` would use.
    ///
    /// Derived from the commands for the same reason [`Self::last_point`] is,
    /// with one difference that matters: `save`/`restore` are part of the
    /// answer. The font is paint STATE, so a `restore()` puts back whatever
    /// was in effect at its `save()`, and a backwards scan for the last
    /// `SetFont` would happily return one that has since been popped. So this
    /// walks FORWARD over a stack, which is what a canvas does.
    pub fn current_font(&self) -> Font {
        let mut font = Font::default();
        let mut saved: Vec<Font> = Vec::new();
        for command in &self.commands {
            match command {
                DrawCmd::SetFont(f) => font = f.clone(),
                DrawCmd::Save => saved.push(font.clone()),
                // An unbalanced `restore()` is a no-op, per spec: it pops
                // nothing rather than trapping.
                DrawCmd::Restore => {
                    if let Some(previous) = saved.pop() {
                        font = previous;
                    }
                }
                _ => {}
            }
        }
        font
    }

    fn last_point(&self) -> Option<(f32, f32)> {
        for command in self.commands.iter().rev() {
            match command {
                DrawCmd::MoveTo(x, y) | DrawCmd::LineTo(x, y) => return Some((*x, *y)),
                DrawCmd::QuadraticCurveTo { x, y, .. }
                | DrawCmd::BezierCurveTo { x, y, .. } => return Some((*x, *y)),
                // An arc ends where its sweep does.
                DrawCmd::Arc { x, y, r, end, .. } => {
                    return Some((x + r * end.cos(), y + r * end.sin()));
                }
                // A rectangle is a closed subpath: it ends where it began.
                DrawCmd::Rect { x, y, .. } => return Some((*x, *y)),
                // `beginPath` discards everything before it, so the search
                // stops rather than reaching into a path that no longer exists.
                DrawCmd::BeginPath => return None,
                _ => {}
            }
        }
        None
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

    fn set_image_smoothing(&mut self, enabled: bool) {
        self.commands.push(DrawCmd::SetImageSmoothing(enabled));
    }
    fn set_text_align(&mut self, align: super::TextAlign) {
        self.commands.push(DrawCmd::SetTextAlign(align));
    }
    fn set_text_baseline(&mut self, baseline: super::TextBaseline) {
        self.commands.push(DrawCmd::SetTextBaseline(baseline));
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
    fn set_fill_paint(&mut self, paint: &super::Paint) {
        self.commands.push(DrawCmd::SetFillPaint(paint.clone()));
    }
    fn set_stroke_paint(&mut self, paint: &super::Paint) {
        self.commands.push(DrawCmd::SetStrokePaint(paint.clone()));
    }
    fn set_shadow(&mut self, shadow: &super::Shadow) {
        self.commands.push(DrawCmd::SetShadow(*shadow));
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
    fn ellipse_arc(
        &mut self,
        x: f32,
        y: f32,
        rx: f32,
        ry: f32,
        rotation: f32,
        start: f32,
        end: f32,
        ccw: bool,
    ) {
        self.commands.push(DrawCmd::EllipseArc {
            x,
            y,
            rx,
            ry,
            rotation,
            start,
            end,
            ccw,
        });
    }

    // Drawing
    fn fill(&mut self) {
        self.commands.push(DrawCmd::Fill);
    }
    fn fill_with_rule(&mut self, rule: super::FillRule) {
        self.commands.push(DrawCmd::FillWithRule(rule));
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
    fn put_image_data(&mut self, img: &Image, dx: f32, dy: f32) {
        self.commands.push(DrawCmd::PutImageData {
            image: img.clone(),
            dx,
            dy,
        });
    }
    #[allow(clippy::too_many_arguments)]
    fn draw_image_rect(
        &mut self,
        img: &Image,
        sx: f32,
        sy: f32,
        sw: f32,
        sh: f32,
        dx: f32,
        dy: f32,
        dw: f32,
        dh: f32,
    ) {
        self.commands.push(DrawCmd::DrawImageRect {
            image: img.clone(),
            sx,
            sy,
            sw,
            sh,
            dx,
            dy,
            dw,
            dh,
        });
    }
    fn clip(&mut self) {
        self.commands.push(DrawCmd::Clip);
    }
    fn clip_with_rule(&mut self, rule: super::FillRule) {
        self.commands.push(DrawCmd::ClipWithRule(rule));
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

    fn current_point(&self) -> Option<(f32, f32)> {
        self.last_point()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Where the recorded path ends, for the arc tests below.
    fn end_of(c: &RecordingCanvas) -> Option<(f32, f32)> {
        c.last_point()
    }

    /// A gradient must PAINT, not merely record. This is the whole point of
    /// `fillStyle` accepting one: before it, the state held a bare `Color` and
    /// a gradient was unexpressible, so `LinearGradientBrush` had nowhere to
    /// land.
    ///
    /// Asserted on PIXELS through the real backend, because a recording
    /// assertion would pass even if the shader were never built.
    #[test]
    fn a_linear_gradient_paints_a_ramp_not_a_flat_fill() {
        use crate::canvas::{Gradient, Paint, TinySkiaCanvas};

        let mut pixmap = tiny_skia::Pixmap::new(64, 8).expect("pixmap");
        let mut gradient = Gradient::linear(0.0, 0.0, 64.0, 0.0);
        gradient.add_color_stop(0.0, Color::rgb(255, 0, 0));
        gradient.add_color_stop(1.0, Color::rgb(0, 0, 255));

        {
            let mut c = TinySkiaCanvas::new(&mut pixmap);
            c.set_fill_paint(&Paint::Gradient(gradient));
            c.fill_rect(0.0, 0.0, 64.0, 8.0);
        }

        let px = |x: u32| {
            let p = pixmap.pixel(x, 4).expect("pixel in bounds");
            (p.red(), p.blue())
        };
        let (left_r, left_b) = px(1);
        let (right_r, right_b) = px(62);

        assert!(
            left_r > 200 && left_b < 60,
            "the left end is the first stop (red), got r={left_r} b={left_b}",
        );
        assert!(
            right_b > 200 && right_r < 60,
            "the right end is the last stop (blue), got r={right_r} b={right_b}",
        );
        // The ramp is the assertion that separates a gradient from a flat
        // fill of either stop: the middle must be neither end.
        let (mid_r, mid_b) = px(32);
        assert!(
            mid_r > 40 && mid_b > 40 && mid_r < 220 && mid_b < 220,
            "the middle interpolates, got r={mid_r} b={mid_b}",
        );
    }

    /// `evenodd` and `nonzero` disagree about a hole, and that disagreement is
    /// the only reason the rule is a parameter. A backend that ignored it
    /// would fill the hole and this would catch it.
    #[test]
    fn the_even_odd_fill_rule_leaves_a_hole_that_nonzero_fills() {
        use crate::canvas::{FillRule, TinySkiaCanvas};

        // Two nested squares wound the SAME way: under `nonzero` the winding
        // numbers add and the inner square is inside; under `evenodd` they
        // cancel and it is a hole.
        let paint_with = |rule: FillRule| {
            let mut pixmap = tiny_skia::Pixmap::new(40, 40).expect("pixmap");
            {
                let mut c = TinySkiaCanvas::new(&mut pixmap);
                c.set_fill_color(Color::rgb(0, 0, 0));
                c.begin_path();
                c.rect(0.0, 0.0, 40.0, 40.0);
                c.rect(10.0, 10.0, 20.0, 20.0);
                c.fill_with_rule(rule);
            }
            pixmap.pixel(20, 20).expect("centre pixel").alpha()
        };

        assert_eq!(paint_with(FillRule::NonZero), 255, "nonzero fills the centre");
        assert_eq!(paint_with(FillRule::EvenOdd), 0, "evenodd punches a hole");
    }

    #[test]
    fn the_current_point_is_derived_from_the_commands() {
        // Derived rather than tracked in a field, so it cannot drift from what
        // was actually recorded.
        let mut c = RecordingCanvas::new();
        assert_eq!(end_of(&c), None, "no subpath yet");
        c.move_to(10.0, 20.0);
        assert_eq!(end_of(&c), Some((10.0, 20.0)));
        c.line_to(30.0, 40.0);
        assert_eq!(end_of(&c), Some((30.0, 40.0)));
        // `beginPath` discards the path, so the search must stop there rather
        // than reach into one that no longer exists.
        c.begin_path();
        assert_eq!(end_of(&c), None);
    }

    #[test]
    fn arc_to_falls_back_to_a_line_in_every_degenerate_case() {
        // These are the SPEC's own rules, not guards: a zero radius or three
        // collinear points is defined to add a straight line to (x1, y1).
        // `roundRect` with a zero radius goes through here on every corner.
        for (radius, label) in [(0.0, "zero radius"), (-5.0, "negative radius")] {
            let mut c = RecordingCanvas::new();
            c.move_to(0.0, 0.0);
            c.arc_to(10.0, 0.0, 10.0, 10.0, radius);
            assert_eq!(end_of(&c), Some((10.0, 0.0)), "{label} is a line to (x1,y1)");
            assert!(
                !c.commands.iter().any(|cmd| matches!(cmd, DrawCmd::Arc { .. })),
                "{label} must not add an arc"
            );
        }
        // Collinear: no corner to fillet.
        let mut c = RecordingCanvas::new();
        c.move_to(0.0, 0.0);
        c.arc_to(10.0, 0.0, 20.0, 0.0, 5.0);
        assert!(!c.commands.iter().any(|cmd| matches!(cmd, DrawCmd::Arc { .. })));
        // With no subpath at all, the spec says start one at (x1, y1).
        let mut c = RecordingCanvas::new();
        c.arc_to(7.0, 8.0, 20.0, 20.0, 5.0);
        assert_eq!(end_of(&c), Some((7.0, 8.0)));
    }

    #[test]
    fn arc_to_fillets_a_right_angle_where_the_tangents_meet() {
        // A 90° corner with radius r: the tangent points sit exactly r back
        // along each leg, which is the one case the geometry can be checked by
        // hand.
        let mut c = RecordingCanvas::new();
        c.move_to(0.0, 0.0);
        c.arc_to(10.0, 0.0, 10.0, 10.0, 4.0);
        // The line runs to the first tangent point, 4 back from the corner.
        assert!(
            c.commands
                .iter()
                .any(|cmd| matches!(cmd, DrawCmd::LineTo(x, y) if (*x - 6.0).abs() < 0.01 && y.abs() < 0.01)),
            "line to the tangent point at (6, 0)"
        );
        let arc = c
            .commands
            .iter()
            .find_map(|cmd| match cmd {
                DrawCmd::Arc { x, y, r, .. } => Some((*x, *y, *r)),
                _ => None,
            })
            .expect("an arc rounds the corner");
        assert!((arc.0 - 6.0).abs() < 0.01 && (arc.1 - 4.0).abs() < 0.01, "centre at (6, 4), got {arc:?}");
        assert!((arc.2 - 4.0).abs() < 0.01);
    }

    #[test]
    fn round_rect_clamps_a_radius_that_would_overlap() {
        // A radius larger than half the shorter side is scaled down rather than
        // allowed to turn the shape inside out — the spec requires it.
        let mut c = RecordingCanvas::new();
        c.round_rect(0.0, 0.0, 20.0, 10.0, 40.0);
        let radii: Vec<f32> = c
            .commands
            .iter()
            .filter_map(|cmd| match cmd {
                DrawCmd::Arc { r, .. } => Some(*r),
                _ => None,
            })
            .collect();
        assert!(!radii.is_empty(), "the corners are arcs");
        assert!(
            radii.iter().all(|r| (*r - 5.0).abs() < 0.01),
            "clamped to half the 10px side, got {radii:?}"
        );
    }

    #[test]
    fn round_rect_with_no_radius_is_a_rectangle_of_straight_lines() {
        let mut c = RecordingCanvas::new();
        c.round_rect(0.0, 0.0, 20.0, 10.0, 0.0);
        assert!(
            !c.commands.iter().any(|cmd| matches!(cmd, DrawCmd::Arc { .. })),
            "no corner to round"
        );
        assert!(c.commands.iter().any(|cmd| matches!(cmd, DrawCmd::ClosePath)));
    }

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

#[cfg(test)]
mod put_image_data_tests {
    use super::*;
    use crate::canvas::Color;

    fn red_pixel() -> Image {
        Image::from_rgba(1, 1, vec![255, 0, 0, 255])
    }

    /// `putImageData` is a RAW write, so it is its own command — not a
    /// `drawImage` sized to the image.
    ///
    /// The two take the same arguments and honour different things: this one
    /// ignores the transform, the clip, `globalAlpha` and the blend mode. Were
    /// it recorded as a `DrawImage`, replay would pick all four up from
    /// whatever state it had reached by then, and a software renderer's frame
    /// would land somewhere else, tinted, or clipped away.
    #[test]
    fn put_image_data_records_as_itself_not_as_a_draw_image() {
        let mut rec = RecordingCanvas::new();
        rec.set_global_alpha(0.5);
        rec.translate(100.0, 100.0);
        rec.put_image_data(&red_pixel(), 7.0, 9.0);

        let written = rec
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCmd::PutImageData { .. }))
            .count();
        assert_eq!(written, 1, "recorded once, as itself");
        assert!(
            !rec.commands
                .iter()
                .any(|c| matches!(c, DrawCmd::DrawImage { .. })),
            "and never as a draw_image"
        );
        // The destination is what the caller said, unshifted by the
        // translate that precedes it — the transform is not applied at
        // record time either.
        assert!(matches!(
            rec.commands.last(),
            Some(DrawCmd::PutImageData { dx, dy, .. }) if *dx == 7.0 && *dy == 9.0
        ));
    }

    /// Replay dispatches it to `put_image_data`, not to `draw_image`.
    #[test]
    fn replay_keeps_the_two_apart() {
        let mut rec = RecordingCanvas::new();
        rec.put_image_data(&red_pixel(), 1.0, 2.0);
        rec.draw_image(&red_pixel(), 3.0, 4.0, 5.0, 6.0);

        let mut replayed = RecordingCanvas::new();
        rec.replay(&mut replayed);
        assert_eq!(replayed.commands, rec.commands, "one arm each, in order");
    }

    /// Paint state is recorded around it and belongs to the OTHER commands —
    /// this is the guard against someone "simplifying" the arm later.
    #[test]
    fn surrounding_state_is_untouched_by_the_raw_write() {
        let mut rec = RecordingCanvas::new();
        rec.set_fill_color(Color { r: 1, g: 2, b: 3, a: 4 });
        rec.put_image_data(&red_pixel(), 0.0, 0.0);
        rec.fill_rect(0.0, 0.0, 10.0, 10.0);
        assert!(matches!(rec.commands.first(), Some(DrawCmd::SetFillColor(_))));
        assert!(matches!(rec.commands.last(), Some(DrawCmd::FillRect { .. })));
    }
}
