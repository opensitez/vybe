//! The 2D canvas, as webcore exposes it.
//!
//! The API IS webcanvas's — the same crate webcore re-exports — so the types
//! are identical by construction. What differs is who paints: each `<canvas>`
//! keeps a [`RecordingCanvas`] here, which answers every state question
//! (`getTransform`, `fillStyle`, `lineWidth`, …) without a round trip, and the
//! commands it records are shipped to the page's real `<canvas>` as one message
//! per [`Document::with_canvas_2d`](crate::types::Document::with_canvas_2d)
//! call. Only the questions that need pixels or font metrics ask the browser.

use std::cell::Cell;
use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};

use base64::Engine as _;
use serde::Deserialize;
use serde_json::{Value, json};

pub use webcanvas::canvas::*;

use crate::page;

/// Parse a CSS colour the way the page's own CSS engine does.
///
/// Asked of the browser once per distinct string and cached: a paint loop that
/// sets the same `fillStyle` every frame asks once.
pub fn parse_color_css(css: &str) -> Option<Color> {
    static CACHE: OnceLock<Mutex<HashMap<String, Option<Color>>>> = OnceLock::new();
    let cache = CACHE.get_or_init(Default::default);
    if let Some(hit) = cache.lock().ok().and_then(|cache| cache.get(css).copied()) {
        return hit;
    }
    let rgba: Option<[u8; 4]> = page::read(0, "canvas_parse_color", json!([css]));
    let color = rgba.map(|[r, g, b, a]| Color { r, g, b, a });
    if let Ok(mut cache) = cache.lock() {
        cache.insert(css.to_string(), color);
    }
    color
}

/// Parse a CSS `font` shorthand the way the page's canvas does. Cached.
fn parse_font_css(css: &str) -> Option<Font> {
    #[derive(Deserialize)]
    struct Wire {
        family: String,
        size: f32,
        bold: bool,
        italic: bool,
    }
    static CACHE: OnceLock<Mutex<HashMap<String, Option<Font>>>> = OnceLock::new();
    let cache = CACHE.get_or_init(Default::default);
    if let Some(hit) = cache.lock().ok().and_then(|cache| cache.get(css).cloned()) {
        return hit;
    }
    let wire: Option<Wire> = page::read(0, "canvas_parse_font", json!([css]));
    let font = wire.map(|w| Font {
        family: w.family,
        size: w.size,
        weight: if w.bold { FontWeight::Bold } else { FontWeight::Normal },
        style: if w.italic { FontStyle::Italic } else { FontStyle::Normal },
    });
    if let Ok(mut cache) = cache.lock() {
        cache.insert(css.to_string(), font.clone());
    }
    font
}

fn browser_syntax() -> CanvasSyntax {
    CanvasSyntax {
        parse_color: parse_color_css,
        parse_font: parse_font_css,
        ..CanvasSyntax::default()
    }
}

/// Each element's recording: its paint state, and the commands not yet sent.
fn contexts() -> &'static Mutex<HashMap<(u64, u32), RecordingCanvas>> {
    static CONTEXTS: OnceLock<Mutex<HashMap<(u64, u32), RecordingCanvas>>> = OnceLock::new();
    CONTEXTS.get_or_init(Default::default)
}

pub(crate) fn reset() {
    if let Ok(mut contexts) = contexts().lock() {
        contexts.clear();
    }
}

/// Run `f` against the 2D context of `node` in `page`, then send what it drew.
pub(crate) fn with_context<R>(page: u64, node: u32, f: impl FnOnce(&mut dyn Canvas) -> R) -> R {
    let mut contexts = contexts().lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    let recording = contexts.entry((page, node)).or_insert_with(|| {
        let mut recording = RecordingCanvas::new();
        recording.set_syntax(browser_syntax());
        recording
    });
    let mut context = RemoteCanvas { page, node, recording, sent: Cell::new(0) };
    let result = f(&mut context);
    context.flush();
    context.recording.commands.clear();
    result
}

struct RemoteCanvas<'a> {
    page: u64,
    node: u32,
    recording: &'a mut RecordingCanvas,
    /// How many of the recording's commands have already gone to the page.
    sent: Cell<usize>,
}

impl RemoteCanvas<'_> {
    /// Send every command recorded since the last send, in order.
    ///
    /// Commands without pixels travel together as one message; an image goes
    /// in its own, with its bytes on the binary channel rather than in JSON.
    fn flush(&self) {
        let commands = &self.recording.commands[self.sent.get()..];
        if commands.is_empty() {
            return;
        }
        let mut batch = Vec::new();
        for command in commands {
            match command_pixels(command) {
                Some((pixels, wire)) => {
                    self.send_batch(&mut batch);
                    page::write_pixels(self.page, json!({ "node": self.node, "command": wire }), pixels);
                }
                None => batch.push(command_json(command)),
            }
        }
        self.send_batch(&mut batch);
        self.sent.set(self.recording.commands.len());
    }

    fn send_batch(&self, batch: &mut Vec<Value>) {
        if !batch.is_empty() {
            page::write(self.page, "canvas_draw", json!([self.node, std::mem::take(batch)]));
        }
    }

    /// Ask the page's context a question, after it has drawn everything so far.
    fn ask<T: serde::de::DeserializeOwned + Default>(&self, question: &str, args: Value) -> T {
        self.flush();
        page::read(self.page, "canvas_query", json!([self.node, question, args]))
    }
}

fn color(c: &Color) -> Value {
    json!([c.r, c.g, c.b, c.a])
}

fn matrix(m: &Matrix) -> Value {
    json!([m.a, m.b, m.c, m.d, m.e, m.f])
}

fn rule(r: &FillRule) -> &'static str {
    match r {
        FillRule::NonZero => "nonzero",
        FillRule::EvenOdd => "evenodd",
    }
}

fn path(p: &Path2D) -> Value {
    Value::Array(
        p.ops
            .iter()
            .map(|op| match *op {
                PathOp::ClosePath => json!(["closePath"]),
                PathOp::MoveTo(x, y) => json!(["moveTo", x, y]),
                PathOp::LineTo(x, y) => json!(["lineTo", x, y]),
                PathOp::QuadraticCurveTo { cx, cy, x, y } => json!(["quadraticCurveTo", cx, cy, x, y]),
                PathOp::BezierCurveTo { cx1, cy1, cx2, cy2, x, y } => {
                    json!(["bezierCurveTo", cx1, cy1, cx2, cy2, x, y])
                }
                PathOp::ArcTo { x1, y1, x2, y2, radius } => json!(["arcTo", x1, y1, x2, y2, radius]),
                PathOp::Rect { x, y, w, h } => json!(["rect", x, y, w, h]),
                PathOp::RoundRect { x, y, w, h, radii } => json!(["roundRect", x, y, w, h, radii]),
                PathOp::Arc { x, y, r, start, end, ccw } => json!(["arc", x, y, r, start, end, ccw]),
                PathOp::Ellipse { x, y, rx, ry, rotation, start, end, ccw } => {
                    json!(["ellipse", x, y, rx, ry, rotation, start, end, ccw])
                }
            })
            .collect(),
    )
}

fn gradient(g: &Gradient) -> Value {
    let stops: Vec<Value> = g.stops.iter().map(|s| json!([s.offset, color(&s.color)])).collect();
    match g.kind {
        GradientKind::Linear { x0, y0, x1, y1 } => json!({ "linear": [x0, y0, x1, y1], "stops": stops }),
        GradientKind::Radial { x0, y0, r0, x1, y1, r1 } => {
            json!({ "radial": [x0, y0, r0, x1, y1, r1], "stops": stops })
        }
        GradientKind::Conic { start_angle, x, y } => json!({ "conic": [start_angle, x, y], "stops": stops }),
    }
}

fn repetition(r: &Repetition) -> &'static str {
    match r {
        Repetition::Repeat => "repeat",
        Repetition::RepeatX => "repeat-x",
        Repetition::RepeatY => "repeat-y",
        Repetition::NoRepeat => "no-repeat",
    }
}

/// The pixels a command carries, if any, and the command without them.
fn command_pixels(command: &DrawCmd) -> Option<(Vec<u8>, Value)> {
    let image = |img: &Image| json!({ "width": img.width, "height": img.height });
    let pattern = |target: &str, p: &Pattern| {
        (
            p.image.pixels.to_vec(),
            json!([target, image(&p.image), repetition(&p.repetition), matrix(&p.transform)]),
        )
    };
    match command {
        DrawCmd::PutImageData { image: img, dx, dy } => {
            Some((img.pixels.to_vec(), json!(["putImageData", image(img), dx, dy])))
        }
        DrawCmd::DrawImage { image: img, x, y, w, h } => {
            Some((img.pixels.to_vec(), json!(["drawImage", image(img), x, y, w, h])))
        }
        DrawCmd::DrawImageRect { image: img, sx, sy, sw, sh, dx, dy, dw, dh } => Some((
            img.pixels.to_vec(),
            json!(["drawImageRect", image(img), sx, sy, sw, sh, dx, dy, dw, dh]),
        )),
        DrawCmd::SetFillPaint(Paint::Pattern(p)) => Some(pattern("fillPattern", p)),
        DrawCmd::SetStrokePaint(Paint::Pattern(p)) => Some(pattern("strokePattern", p)),
        _ => None,
    }
}

/// A recorded call as the page's replay reads it: `[method, ...args]`.
fn command_json(command: &DrawCmd) -> Value {
    use DrawCmd as C;
    match command {
        C::SetFillColor(c) => json!(["fillColor", color(c)]),
        C::SetStrokeColor(c) => json!(["strokeColor", color(c)]),
        C::SetLineWidth(w) => json!(["set", "lineWidth", w]),
        C::SetLineCap(cap) => json!(["set", "lineCap", cap.as_str()]),
        C::SetLineJoin(join) => json!(["set", "lineJoin", join.as_str()]),
        C::SetMiterLimit(limit) => json!(["set", "miterLimit", limit]),
        C::SetGlobalAlpha(alpha) => json!(["set", "globalAlpha", alpha]),
        C::SetImageSmoothing(on) => json!(["set", "imageSmoothingEnabled", on]),
        C::SetTextAlign(align) => json!(["set", "textAlign", align.as_str()]),
        C::SetTextBaseline(baseline) => json!(["set", "textBaseline", baseline.as_str()]),
        C::SetFont(font) => json!(["set", "font", font.to_css()]),
        C::SetLineDash(dash) => json!(["setLineDash", dash]),
        C::SetLineDashOffset(offset) => json!(["set", "lineDashOffset", offset]),
        C::SetFillPaint(Paint::Color(c)) => json!(["fillColor", color(c)]),
        C::SetStrokePaint(Paint::Color(c)) => json!(["strokeColor", color(c)]),
        C::SetFillPaint(Paint::Gradient(g)) => json!(["fillGradient", gradient(g)]),
        C::SetStrokePaint(Paint::Gradient(g)) => json!(["strokeGradient", gradient(g)]),
        // Patterns carry pixels and are sent by `command_pixels`.
        C::SetFillPaint(Paint::Pattern(_)) | C::SetStrokePaint(Paint::Pattern(_)) => json!(["noop"]),
        C::SetShadow(s) => json!(["shadow", color(&s.color), s.blur, s.offset_x, s.offset_y]),
        C::BeginPath => json!(["beginPath"]),
        C::ClosePath => json!(["closePath"]),
        C::MoveTo(x, y) => json!(["moveTo", x, y]),
        C::LineTo(x, y) => json!(["lineTo", x, y]),
        C::QuadraticCurveTo { cx, cy, x, y } => json!(["quadraticCurveTo", cx, cy, x, y]),
        C::BezierCurveTo { cx1, cy1, cx2, cy2, x, y } => json!(["bezierCurveTo", cx1, cy1, cx2, cy2, x, y]),
        C::Arc { x, y, r, start, end, ccw } => json!(["arc", x, y, r, start, end, ccw]),
        C::Rect { x, y, w, h } => json!(["rect", x, y, w, h]),
        C::Ellipse { x, y, rx, ry } => json!(["ellipse", x, y, rx, ry, 0.0, 0.0, std::f32::consts::TAU, false]),
        C::EllipseArc { x, y, rx, ry, rotation, start, end, ccw } => {
            json!(["ellipse", x, y, rx, ry, rotation, start, end, ccw])
        }
        C::Fill => json!(["fill"]),
        C::FillWithRule(r) => json!(["fill", rule(r)]),
        C::Stroke => json!(["stroke"]),
        C::FillRect { x, y, w, h } => json!(["fillRect", x, y, w, h]),
        C::StrokeRect { x, y, w, h } => json!(["strokeRect", x, y, w, h]),
        C::ClearRect { x, y, w, h } => json!(["clearRect", x, y, w, h]),
        C::FillText { text, x, y } => json!(["fillText", text, x, y]),
        C::StrokeText { text, x, y } => json!(["strokeText", text, x, y]),
        // Images carry pixels and are sent by `command_pixels`.
        C::PutImageData { .. } | C::DrawImage { .. } | C::DrawImageRect { .. } => json!(["noop"]),
        C::Clip => json!(["clip"]),
        C::ClipWithRule(r) => json!(["clip", rule(r)]),
        C::ResetClip => json!(["resetClip"]),
        C::Save => json!(["save"]),
        C::Restore => json!(["restore"]),
        C::Translate(x, y) => json!(["translate", x, y]),
        C::Rotate(angle) => json!(["rotate", angle]),
        C::Scale(x, y) => json!(["scale", x, y]),
        C::Transform { m11, m12, m21, m22, dx, dy } => json!(["transform", m11, m12, m21, m22, dx, dy]),
        C::ResetTransform => json!(["resetTransform"]),
        C::SetTransform(m) => json!(["setTransform", matrix(m)]),
        C::Reset => json!(["reset"]),
        C::SetGlobalCompositeOperation(op) => json!(["set", "globalCompositeOperation", op.as_str()]),
        C::SetImageSmoothingQuality(q) => json!(["set", "imageSmoothingQuality", q.as_str()]),
        C::SetFilter(filter) => json!(["set", "filter", filter]),
        C::SetDirection(d) => json!(["set", "direction", d.as_str()]),
        C::SetLang(lang) => json!(["set", "lang", lang]),
        C::SetLetterSpacing(s) => json!(["set", "letterSpacing", s]),
        C::SetWordSpacing(s) => json!(["set", "wordSpacing", s]),
        C::SetFontKerning(k) => json!(["set", "fontKerning", k.as_str()]),
        C::SetFontStretch(s) => json!(["set", "fontStretch", s.as_str()]),
        C::SetFontVariantCaps(c) => json!(["set", "fontVariantCaps", c.as_str()]),
        C::SetTextRendering(r) => json!(["set", "textRendering", r.as_str()]),
        C::DrawFocusIfNeeded(focused) => json!(["drawFocusIfNeeded", focused]),
    }
}

#[derive(Deserialize, Default)]
struct Metrics {
    width: f32,
    actual_bounding_box_left: f32,
    actual_bounding_box_right: f32,
    font_bounding_box_ascent: f32,
    font_bounding_box_descent: f32,
    actual_bounding_box_ascent: f32,
    actual_bounding_box_descent: f32,
    em_height_ascent: f32,
    em_height_descent: f32,
    hanging_baseline: f32,
    alphabetic_baseline: f32,
    ideographic_baseline: f32,
}

#[derive(Deserialize, Default)]
struct Pixels {
    width: u32,
    height: u32,
    data: String,
}

impl Canvas for RemoteCanvas<'_> {
    fn drawing_state(&self) -> &DrawingState {
        self.recording.drawing_state()
    }

    fn set_fill_style_css(&mut self, css: &str) {
        self.recording.set_fill_style_css(css)
    }

    fn set_stroke_style_css(&mut self, css: &str) {
        self.recording.set_stroke_style_css(css)
    }

    fn set_font_css(&mut self, css: &str) {
        self.recording.set_font_css(css)
    }

    fn set_shadow_color_css(&mut self, css: &str) {
        self.recording.set_shadow_color_css(css)
    }

    fn set_fill_color(&mut self, color: Color) {
        self.recording.set_fill_color(color)
    }

    fn set_stroke_color(&mut self, color: Color) {
        self.recording.set_stroke_color(color)
    }

    fn set_line_width(&mut self, width: f32) {
        self.recording.set_line_width(width)
    }

    fn set_line_cap(&mut self, cap: LineCap) {
        self.recording.set_line_cap(cap)
    }

    fn set_line_join(&mut self, join: LineJoin) {
        self.recording.set_line_join(join)
    }

    fn set_miter_limit(&mut self, limit: f32) {
        self.recording.set_miter_limit(limit)
    }

    fn set_global_alpha(&mut self, alpha: f32) {
        self.recording.set_global_alpha(alpha)
    }

    fn set_image_smoothing(&mut self, enabled: bool) {
        self.recording.set_image_smoothing(enabled)
    }

    fn set_text_align(&mut self, align: TextAlign) {
        self.recording.set_text_align(align)
    }

    fn set_text_baseline(&mut self, baseline: TextBaseline) {
        self.recording.set_text_baseline(baseline)
    }

    fn set_font(&mut self, font: &Font) {
        self.recording.set_font(font)
    }

    fn set_line_dash(&mut self, intervals: &[f32]) {
        self.recording.set_line_dash(intervals)
    }

    fn set_line_dash_offset(&mut self, offset: f32) {
        self.recording.set_line_dash_offset(offset)
    }

    fn set_fill_paint(&mut self, paint: &Paint) {
        self.recording.set_fill_paint(paint)
    }

    fn set_stroke_paint(&mut self, paint: &Paint) {
        self.recording.set_stroke_paint(paint)
    }

    fn set_shadow(&mut self, shadow: &Shadow) {
        self.recording.set_shadow(shadow)
    }

    fn begin_path(&mut self) {
        self.recording.begin_path()
    }

    fn close_path(&mut self) {
        self.recording.close_path()
    }

    fn move_to(&mut self, x: f32, y: f32) {
        self.recording.move_to(x, y)
    }

    fn line_to(&mut self, x: f32, y: f32) {
        self.recording.line_to(x, y)
    }

    fn quadratic_curve_to(&mut self, cx: f32, cy: f32, x: f32, y: f32) {
        self.recording.quadratic_curve_to(cx, cy, x, y)
    }

    fn bezier_curve_to(&mut self, cx1: f32, cy1: f32, cx2: f32, cy2: f32, x: f32, y: f32) {
        self.recording.bezier_curve_to(cx1, cy1, cx2, cy2, x, y)
    }

    fn arc(&mut self, x: f32, y: f32, r: f32, start: f32, end: f32, ccw: bool) {
        self.recording.arc(x, y, r, start, end, ccw)
    }

    fn rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.recording.rect(x, y, w, h)
    }

    fn ellipse(&mut self, x: f32, y: f32, rx: f32, ry: f32) {
        self.recording.ellipse(x, y, rx, ry)
    }

    fn ellipse_arc( &mut self, x: f32, y: f32, rx: f32, ry: f32, rotation: f32, start: f32, end: f32, ccw: bool, ) {
        self.recording.ellipse_arc(x, y, rx, ry, rotation, start, end, ccw)
    }

    fn fill(&mut self) {
        self.recording.fill()
    }

    fn fill_with_rule(&mut self, rule: FillRule) {
        self.recording.fill_with_rule(rule)
    }

    fn stroke(&mut self) {
        self.recording.stroke()
    }

    fn fill_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.recording.fill_rect(x, y, w, h)
    }

    fn stroke_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.recording.stroke_rect(x, y, w, h)
    }

    fn clear_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.recording.clear_rect(x, y, w, h)
    }

    fn fill_text(&mut self, text: &str, x: f32, y: f32) {
        self.recording.fill_text(text, x, y)
    }

    fn stroke_text(&mut self, text: &str, x: f32, y: f32) {
        self.recording.stroke_text(text, x, y)
    }

    fn draw_image(&mut self, img: &Image, x: f32, y: f32, w: f32, h: f32) {
        self.recording.draw_image(img, x, y, w, h)
    }

    fn put_image_data(&mut self, img: &Image, dx: f32, dy: f32) {
        self.recording.put_image_data(img, dx, dy)
    }

    fn draw_image_rect( &mut self, img: &Image, sx: f32, sy: f32, sw: f32, sh: f32, dx: f32, dy: f32, dw: f32, dh: f32, ) {
        self.recording.draw_image_rect(img, sx, sy, sw, sh, dx, dy, dw, dh)
    }

    fn clip(&mut self) {
        self.recording.clip()
    }

    fn clip_with_rule(&mut self, rule: FillRule) {
        self.recording.clip_with_rule(rule)
    }

    fn reset_clip(&mut self) {
        self.recording.reset_clip()
    }

    fn save(&mut self) {
        self.recording.save()
    }

    fn restore(&mut self) {
        self.recording.restore()
    }

    fn translate(&mut self, x: f32, y: f32) {
        self.recording.translate(x, y)
    }

    fn rotate(&mut self, rad: f32) {
        self.recording.rotate(rad)
    }

    fn scale(&mut self, sx: f32, sy: f32) {
        self.recording.scale(sx, sy)
    }

    fn transform(&mut self, m11: f32, m12: f32, m21: f32, m22: f32, dx: f32, dy: f32) {
        self.recording.transform(m11, m12, m21, m22, dx, dy)
    }

    fn reset_transform(&mut self) {
        self.recording.reset_transform()
    }

    fn set_transform(&mut self, m: Matrix) {
        self.recording.set_transform(m)
    }

    fn get_transform(&self) -> Matrix {
        self.recording.get_transform()
    }

    fn current_point(&self) -> Option<(f32, f32)> {
        self.recording.current_point()
    }

    fn reset(&mut self) {
        self.recording.reset()
    }

    fn set_global_composite_operation(&mut self, op: CompositeOp) {
        self.recording.set_global_composite_operation(op)
    }

    fn set_image_smoothing_quality(&mut self, quality: SmoothingQuality) {
        self.recording.set_image_smoothing_quality(quality)
    }

    fn set_filter(&mut self, filter: &str) {
        self.recording.set_filter(filter)
    }

    fn get_line_dash(&self) -> Vec<f32> {
        self.recording.get_line_dash()
    }

    fn set_direction(&mut self, direction: Direction) {
        self.recording.set_direction(direction)
    }

    fn set_lang(&mut self, lang: &str) {
        self.recording.set_lang(lang)
    }

    fn set_letter_spacing(&mut self, spacing: &str) {
        self.recording.set_letter_spacing(spacing)
    }

    fn set_word_spacing(&mut self, spacing: &str) {
        self.recording.set_word_spacing(spacing)
    }

    fn set_font_kerning(&mut self, kerning: FontKerning) {
        self.recording.set_font_kerning(kerning)
    }

    fn set_font_stretch(&mut self, stretch: FontStretch) {
        self.recording.set_font_stretch(stretch)
    }

    fn set_font_variant_caps(&mut self, caps: FontVariantCaps) {
        self.recording.set_font_variant_caps(caps)
    }

    fn set_text_rendering(&mut self, rendering: TextRendering) {
        self.recording.set_text_rendering(rendering)
    }

    fn draw_focus_if_needed(&mut self, focused: bool) {
        self.recording.draw_focus_if_needed(focused)
    }

    // ── The questions only the page can answer ──

    fn measure_text(&mut self, text: &str) -> TextMetrics {
        let m: Metrics = self.ask("measureText", json!([text]));
        TextMetrics {
            width: m.width,
            actual_bounding_box_left: m.actual_bounding_box_left,
            actual_bounding_box_right: m.actual_bounding_box_right,
            font_bounding_box_ascent: m.font_bounding_box_ascent,
            font_bounding_box_descent: m.font_bounding_box_descent,
            actual_bounding_box_ascent: m.actual_bounding_box_ascent,
            actual_bounding_box_descent: m.actual_bounding_box_descent,
            em_height_ascent: m.em_height_ascent,
            em_height_descent: m.em_height_descent,
            hanging_baseline: m.hanging_baseline,
            alphabetic_baseline: m.alphabetic_baseline,
            ideographic_baseline: m.ideographic_baseline,
        }
    }

    fn is_point_in_path(&self, x: f32, y: f32, fill_rule: FillRule) -> bool {
        self.ask("isPointInPath", json!([x, y, rule(&fill_rule)]))
    }

    fn is_point_in_stroke(&self, x: f32, y: f32) -> bool {
        self.ask("isPointInStroke", json!([x, y]))
    }

    fn is_point_in_path2d(&self, p: &Path2D, x: f32, y: f32, fill_rule: FillRule) -> bool {
        self.ask("isPointInPath2D", json!([path(p), x, y, rule(&fill_rule)]))
    }

    fn is_point_in_stroke2d(&self, p: &Path2D, x: f32, y: f32) -> bool {
        self.ask("isPointInStroke2D", json!([path(p), x, y]))
    }

    fn get_image_data(&self, sx: i32, sy: i32, sw: u32, sh: u32) -> Option<ImageData> {
        let pixels: Option<Pixels> = self.ask("getImageData", json!([sx, sy, sw, sh]));
        let pixels = pixels?;
        let data = base64::engine::general_purpose::STANDARD.decode(pixels.data).ok()?;
        ImageData::from_rgba(pixels.width, pixels.height, data)
    }

    fn to_blob(&self, mime: &str, quality: Option<f32>) -> Option<Vec<u8>> {
        let encoded: Option<String> = self.ask("toBlob", json!([mime, quality]));
        base64::engine::general_purpose::STANDARD.decode(encoded?).ok()
    }

    fn is_context_lost(&self) -> bool {
        self.ask("isContextLost", json!([]))
    }
}

/// The node a context paints, for diagnostics.
impl std::fmt::Debug for RemoteCanvas<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "RemoteCanvas(page {}, node {})", self.page, self.node)
    }
}
