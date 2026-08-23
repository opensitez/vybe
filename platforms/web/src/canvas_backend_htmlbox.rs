//! The painter behind `web:canvas` when htmlbox is the engine.
//!
//! The same file as `canvas_backend_widgets`, pointed at the other engine, and
//! that is the whole point of the seam: `getContext(element, "2d")` binds a
//! context to a NODE (HTML §4.12.5), each engine turns its own nodes into its
//! own pixels, and nothing above this layer learns which one is installed.
//!
//! **What this forwards to is a real bitmap.** htmlbox's `<canvas>` keeps its
//! pixels on the element and draws into them immediately, so an op that arrives
//! here has been painted by the time the call returns. That is what lets the
//! asking half of the API mean anything: `measureText` reads the font actually
//! in effect on that element's context, and the queries the seam has yet to
//! carry — `getImageData`, `isPointInPath`, `toDataURL` — have real pixels
//! waiting behind them rather than a list of commands nobody has replayed.

use std::sync::Arc;

use crate::canvas_backend::{
    self, CanvasBackend, GradientDef, GradientKind as SeamGradientKind, Op2D, PathDef,
    PathOp2D, PatternDef, Query2D, Query2DValue, StringAttribute, TextMetrics2D,
};
use rhtmledit::canvas::{
    Canvas as _, Color, ColorStop, CompositeOp, Direction, FillRule, Font, FontKerning,
    FontStretch, FontStyle, FontVariantCaps, FontWeight, Gradient as CanvasGradient, Image,
    ImageData, LineCap, LineJoin, Paint as CanvasPaint, Pattern as CanvasPattern, Repetition,
    Path2D as CanvasPath, PathOp as EnginePathOp, Shadow, SmoothingQuality, TextAlign,
    TextBaseline, TextRendering,
};
use rhtmledit::types::Document;

struct HtmlBoxBackend;

fn color(r: u8, g: u8, b: u8, a: u8) -> Color {
    Color { r, g, b, a }
}

/// The node a target names.
///
/// Two forms, in the order a caller means them — the same two
/// `canvas_backend_widgets` resolves, because they are the seam's forms and not
/// an engine's:
///
/// 1. `n<id>` — what an element-bound context carries. `getContext` derives it
///    from the node it was given, so this is the direct case and no search
///    happens.
/// 2. A control NAME — .NET `CreateGraphics` and Flutter's canvas bridge still
///    pass one. MIGRATION ONLY: resolved against `id`, then the `name`
///    attribute, which is what those callers set.
fn node_of(document: &Document, target: &str) -> Option<u32> {
    if let Some(rest) = target.strip_prefix('n') {
        if let Ok(id) = rest.parse::<u32>() {
            return Some(id);
        }
    }
    document
        .get_element_by_id(target)
        .or_else(|| document.query_selector(&format!("[name=\"{target}\"]")))
}

/// Borrow the 2D context `target` names, in the ambient document.
fn with_canvas<T>(target: &str, f: impl FnOnce(&mut dyn rhtmledit::canvas::Canvas) -> T) -> Option<T> {
    crate::engine_htmlbox::with_document(crate::html::active_document(), |document| {
        let node = node_of(document, target)?;
        document.with_canvas_2d(node, f)
    })
    .flatten()
}

impl CanvasBackend for HtmlBoxBackend {
    /// `getContext`'s side effect. The `<canvas>` element owns its bitmap, so
    /// this allocates it if the element has never had one — an element from
    /// `createElement` has not been through the parser.
    fn ensure(&self, target: &str) {
        crate::engine_htmlbox::with_document(crate::html::active_document(), |document| {
            if let Some(node) = node_of(document, target) {
                document.get_context_2d(node);
            }
        });
    }

    /// Answer a question about the surface `target` names.
    ///
    /// Every arm goes through the SAME retained context the paint ops go
    /// through, because every question is about what those ops established:
    /// `getTransform` is about the transforms, `isPointInPath` about the path,
    /// `measureText` about the font. A query answered from anywhere else would
    /// be answering about a different canvas.
    fn query(&self, target: &str, q: Query2D) -> Query2DValue {
        use StringAttribute as A;
        let answer = with_canvas(target, |c| match q {
            Query2D::MeasureText(text) => {
                let m = c.measure_text(&text);
                Query2DValue::Metrics(TextMetrics2D {
                    width: m.width,
                    actual_bounding_box_left: m.actual_bounding_box_left,
                    actual_bounding_box_right: m.actual_bounding_box_right,
                    actual_bounding_box_ascent: m.actual_bounding_box_ascent,
                    actual_bounding_box_descent: m.actual_bounding_box_descent,
                    font_bounding_box_ascent: m.font_bounding_box_ascent,
                    font_bounding_box_descent: m.font_bounding_box_descent,
                    em_height_ascent: m.em_height_ascent,
                    em_height_descent: m.em_height_descent,
                    hanging_baseline: m.hanging_baseline,
                    alphabetic_baseline: m.alphabetic_baseline,
                    ideographic_baseline: m.ideographic_baseline,
                })
            }
            Query2D::GetImageData { sx, sy, sw, sh } => match c.get_image_data(sx, sy, sw, sh) {
                Some(d) => Query2DValue::Pixels {
                    data: d.data,
                    width: d.width,
                    height: d.height,
                },
                None => Query2DValue::Absent,
            },
            Query2D::IsPointInPath { x, y, rule } => Query2DValue::Bool(
                c.is_point_in_path(x, y, FillRule::parse(&rule).unwrap_or(FillRule::NonZero)),
            ),
            Query2D::IsPointInStroke { x, y } => Query2DValue::Bool(c.is_point_in_stroke(x, y)),
            Query2D::GetTransform => {
                let m = c.get_transform();
                Query2DValue::Matrix([m.a, m.b, m.c, m.d, m.e, m.f])
            }
            Query2D::GetLineDash => Query2DValue::Floats(c.get_line_dash()),
            Query2D::IsContextLost => Query2DValue::Bool(c.is_context_lost()),
            Query2D::ToDataUrl { mime, quality } => {
                Query2DValue::Text(c.to_data_url(&mime, quality))
            }
            Query2D::ToBlob { mime, quality } => match c.to_blob(&mime, quality) {
                Some(bytes) => Query2DValue::Bytes(bytes),
                // An unsupported MIME type is `None` from the engine and
                // `Absent` here — the spec says a bad type falls back to PNG at
                // the API layer, which is `canvas.rs`'s job, not a painter's.
                None => Query2DValue::Absent,
            },
            Query2D::IsPointInPathOf { path, x, y, rule } => Query2DValue::Bool(
                c.is_point_in_path2d(
                    &path_of(&path),
                    x,
                    y,
                    FillRule::parse(&rule).unwrap_or(FillRule::NonZero),
                ),
            ),
            Query2D::IsPointInStrokeOf { path, x, y } => {
                Query2DValue::Bool(c.is_point_in_stroke2d(&path_of(&path), x, y))
            }
            Query2D::GetContextAttributes => {
                let a = c.context_attributes();
                Query2DValue::ContextAttributes {
                    alpha: a.alpha,
                    desynchronized: a.desynchronized,
                    color_space: a.color_space.to_string(),
                    color_type: a.color_type.to_string(),
                    will_read_frequently: a.will_read_frequently,
                }
            }
            Query2D::GetStringAttribute(which) => Query2DValue::Text(match which {
                A::Font => c.current_font_css(),
                A::FillStyle => c.fill_style_css(),
                A::StrokeStyle => c.stroke_style_css(),
                A::Filter => c.filter_css(),
                A::GlobalCompositeOperation => c.global_composite_operation().as_str().to_string(),
                A::ImageSmoothingQuality => c.image_smoothing_quality().as_str().to_string(),
                A::ShadowColor => c.shadow_color_css(),
                A::Direction => c.direction().as_str().to_string(),
                A::LetterSpacing => c.letter_spacing_css(),
                A::WordSpacing => c.word_spacing_css(),
                A::FontKerning => c.font_kerning().as_str().to_string(),
                A::FontStretch => c.font_stretch().as_str().to_string(),
                A::FontVariantCaps => c.font_variant_caps().as_str().to_string(),
                A::TextRendering => c.text_rendering().as_str().to_string(),
                A::Lang => c.lang().to_string(),
                A::TextAlign => c.text_align().as_str().to_string(),
                A::TextBaseline => c.text_baseline().as_str().to_string(),
                A::LineCap => c.line_cap().as_str().to_string(),
                A::LineJoin => c.line_join().as_str().to_string(),
            }),
        });
        answer.unwrap_or(Query2DValue::Absent)
    }

    /// Back to a transparent bitmap and a default drawing state — HTML
    /// §4.12.5's `reset()`, which is what "drop everything drawn here" means
    /// for a canvas that paints immediately.
    fn clear_all(&self, target: &str) {
        let _ = with_canvas(target, |c| c.reset());
    }

    fn apply(&self, target: &str, op: Op2D) {
        let _ = with_canvas(target, |c| {
            match op {
                Op2D::Save => c.save(),
                Op2D::Restore => c.restore(),
                Op2D::SetFillStyle(r, g, b, a) => c.set_fill_color(color(r, g, b, a)),
                Op2D::SetStrokeStyle(r, g, b, a) => c.set_stroke_color(color(r, g, b, a)),
                Op2D::SetLineWidth(w) => c.set_line_width(w),
                Op2D::SetLineDash(d) => c.set_line_dash(&d),
                Op2D::SetLineCap(k) => c.set_line_cap(match k.as_str() {
                    "round" => LineCap::Round,
                    "square" => LineCap::Square,
                    _ => LineCap::Butt,
                }),
                Op2D::SetLineJoin(k) => c.set_line_join(match k.as_str() {
                    "round" => LineJoin::Round,
                    "bevel" => LineJoin::Bevel,
                    _ => LineJoin::Miter,
                }),
                Op2D::SetGlobalAlpha(a) => c.set_global_alpha(a),
                Op2D::SetImageSmoothing(on) => c.set_image_smoothing(on),
                Op2D::SetFont {
                    family,
                    size,
                    bold,
                    italic,
                } => c.set_font(&Font {
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
                    },
                }),
                Op2D::Translate(x, y) => c.translate(x, y),
                Op2D::Scale(x, y) => c.scale(x, y),
                Op2D::Rotate(a) => c.rotate(a),
                Op2D::BeginPath => c.begin_path(),
                Op2D::ClosePath => c.close_path(),
                Op2D::MoveTo(x, y) => c.move_to(x, y),
                Op2D::LineTo(x, y) => c.line_to(x, y),
                Op2D::Arc(x, y, r, s, e, ccw) => c.arc(x, y, r, s, e, ccw),
                Op2D::BezierCurveTo(a, b, cc, d, e, f) => c.bezier_curve_to(a, b, cc, d, e, f),
                Op2D::QuadraticCurveTo(a, b, cc, d) => c.quadratic_curve_to(a, b, cc, d),
                Op2D::Rect(x, y, w, h) => c.rect(x, y, w, h),
                // The seam's `Ellipse` carries four arguments where the IDL has
                // eight: rotation, the two angles and the direction are not on
                // it. The engine implements the full one (`ellipse_arc`); this
                // arm reaches the truncated version because that is all the op
                // says.
                Op2D::Ellipse(x, y, rx, ry) => c.ellipse(x, y, rx, ry),
                // `setTransform` is `ResetTransform` then `Transform`, emitted
                // as two ops — composing and replacing are different verbs.
                Op2D::Transform(a, b, cc, d, e, f) => c.transform(a, b, cc, d, e, f),
                Op2D::ResetTransform => c.reset_transform(),
                Op2D::SetMiterLimit(limit) => c.set_miter_limit(limit),
                Op2D::SetLineDashOffset(offset) => c.set_line_dash_offset(offset),
                // An unrecognised keyword leaves the current value, which is
                // what a browser does with a bad assignment to either of these
                // — it is not an error and it is not a reset.
                Op2D::SetTextAlign(k) => {
                    if let Some(a) = TextAlign::parse(&k) {
                        c.set_text_align(a);
                    }
                }
                Op2D::SetTextBaseline(k) => {
                    if let Some(b) = TextBaseline::parse(&k) {
                        c.set_text_baseline(b);
                    }
                }
                Op2D::Fill => c.fill(),
                Op2D::Stroke => c.stroke(),
                Op2D::Clip => c.clip(),
                Op2D::FillRect(x, y, w, h) => c.fill_rect(x, y, w, h),
                Op2D::StrokeRect(x, y, w, h) => c.stroke_rect(x, y, w, h),
                Op2D::ClearRect(x, y, w, h) => c.clear_rect(x, y, w, h),
                Op2D::FillText(t, x, y) => c.fill_text(&t, x, y),
                Op2D::StrokeText(t, x, y) => c.stroke_text(&t, x, y),
                Op2D::PutImageData {
                    pixels,
                    width,
                    height,
                    dx,
                    dy,
                } => {
                    let img = Image::from_rgba(width, height, pixels);
                    c.put_image_data(&img, dx, dy);
                }
                Op2D::DrawImageRgba {
                    pixels,
                    width,
                    height,
                    dx,
                    dy,
                    dw,
                    dh,
                } => {
                    let img = Image::from_rgba(width, height, pixels);
                    c.draw_image(&img, dx, dy, dw, dh);
                }
                Op2D::DrawImagePaletted {
                    indices,
                    palette,
                    width,
                    height,
                    dx,
                    dy,
                    dw,
                    dh,
                } => {
                    // The palette arrives as RGB triples (SDL's shape); the
                    // engine wants packed 0xRRGGBB entries.
                    let packed: Vec<u32> = palette
                        .chunks(3)
                        .map(|c| {
                            ((*c.first().unwrap_or(&0) as u32) << 16)
                                | ((*c.get(1).unwrap_or(&0) as u32) << 8)
                                | (*c.get(2).unwrap_or(&0) as u32)
                        })
                        .collect();
                    let img = Image::from_paletted(width, height, &indices, &packed);
                    c.draw_image(&img, dx, dy, dw, dh);
                }

                // ── CSS values, parsed by the engine ──────────────────
                Op2D::SetFillStyleCss(css) => c.set_fill_style_css(&css),
                Op2D::SetStrokeStyleCss(css) => c.set_stroke_style_css(&css),
                Op2D::SetFontCss(css) => c.set_font_css(&css),
                Op2D::SetShadowColor(css) => c.set_shadow_color_css(&css),
                Op2D::SetFilter(css) => c.set_filter(&css),
                // An unrecognised keyword leaves the attribute as it was — the
                // spec's rule for every enumerated attribute, and the reason
                // these parse rather than default.
                Op2D::SetGlobalCompositeOperation(k) => {
                    if let Some(op) = CompositeOp::parse(&k) {
                        c.set_global_composite_operation(op);
                    }
                }
                Op2D::SetImageSmoothingQuality(k) => {
                    if let Some(q) = SmoothingQuality::parse(&k) {
                        c.set_image_smoothing_quality(q);
                    }
                }

                // ── Shadows ───────────────────────────────────────────
                //
                // Read-modify-write, because the seam sets one component at a
                // time (a page assigns `shadowBlur` without touching
                // `shadowColor`) while the engine holds them as one value.
                Op2D::SetShadowBlur(v) => {
                    let s = Shadow { blur: v, ..current_shadow(c) };
                    c.set_shadow(&s);
                }
                Op2D::SetShadowOffsetX(v) => {
                    let s = Shadow { offset_x: v, ..current_shadow(c) };
                    c.set_shadow(&s);
                }
                Op2D::SetShadowOffsetY(v) => {
                    let s = Shadow { offset_y: v, ..current_shadow(c) };
                    c.set_shadow(&s);
                }

                // ── Text style ────────────────────────────────────────
                Op2D::SetDirection(k) => {
                    if let Some(d) = Direction::parse(&k) {
                        c.set_direction(d);
                    }
                }
                Op2D::SetLetterSpacing(v) => c.set_letter_spacing(&v),
                Op2D::SetWordSpacing(v) => c.set_word_spacing(&v),
                Op2D::SetFontKerning(k) => {
                    if let Some(v) = FontKerning::parse(&k) {
                        c.set_font_kerning(v);
                    }
                }
                Op2D::SetFontStretch(k) => {
                    if let Some(v) = FontStretch::parse(&k) {
                        c.set_font_stretch(v);
                    }
                }
                Op2D::SetFontVariantCaps(k) => {
                    if let Some(v) = FontVariantCaps::parse(&k) {
                        c.set_font_variant_caps(v);
                    }
                }
                Op2D::SetTextRendering(k) => {
                    if let Some(v) = TextRendering::parse(&k) {
                        c.set_text_rendering(v);
                    }
                }
                Op2D::SetLang(v) => c.set_lang(&v),

                // ── Paths ─────────────────────────────────────────────
                Op2D::ArcTo(x1, y1, x2, y2, r) => c.arc_to(x1, y1, x2, y2, r),
                Op2D::RoundRect { x, y, w, h, radii } => c.round_rect_radii(x, y, w, h, radii),
                Op2D::EllipseFull { x, y, rx, ry, rotation, start, end, ccw } => {
                    c.ellipse_arc(x, y, rx, ry, rotation, start, end, ccw)
                }
                Op2D::FillWithRule(rule) => {
                    c.fill_with_rule(FillRule::parse(&rule).unwrap_or(FillRule::NonZero))
                }
                Op2D::ClipWithRule(rule) => {
                    c.clip_with_rule(FillRule::parse(&rule).unwrap_or(FillRule::NonZero))
                }

                // ── Text ──────────────────────────────────────────────
                Op2D::FillTextMaxWidth(t, x, y, max) => {
                    c.fill_text_constrained(&t, x, y, max)
                }
                Op2D::StrokeTextMaxWidth(t, x, y, max) => {
                    c.stroke_text_constrained(&t, x, y, max)
                }

                // ── Gradients and patterns ────────────────────────────
                Op2D::SetFillGradient(def) => {
                    c.set_fill_paint(&CanvasPaint::Gradient(gradient(&def)))
                }
                Op2D::SetStrokeGradient(def) => {
                    c.set_stroke_paint(&CanvasPaint::Gradient(gradient(&def)))
                }
                Op2D::SetFillPattern(def) => {
                    c.set_fill_paint(&CanvasPaint::Pattern(pattern(def)))
                }
                Op2D::SetStrokePattern(def) => {
                    c.set_stroke_paint(&CanvasPaint::Pattern(pattern(def)))
                }

                // ── The rest ──────────────────────────────────────────
                Op2D::Reset => c.reset(),
                Op2D::PutImageDataDirty {
                    pixels, width, height, dx, dy, dirty_x, dirty_y, dirty_w, dirty_h,
                } => {
                    let data = ImageData {
                        width,
                        height,
                        data: pixels,
                        color_space: "srgb",
                    };
                    c.put_image_data_dirty(
                        &data, dx, dy,
                        dirty_x as i32, dirty_y as i32, dirty_w as i32, dirty_h as i32,
                    );
                }
                Op2D::DrawFocusIfNeeded(focused) => c.draw_focus_if_needed(focused),

                // ── Path2D ────────────────────────────────────────────
                //
                // These take the path GIVEN and leave the context's current
                // path untouched, which is the whole difference between
                // `fill(path)` and `fill()`.
                Op2D::FillPath(def, rule) => c.fill_path(
                    &path_of(&def),
                    FillRule::parse(&rule).unwrap_or(FillRule::NonZero),
                ),
                Op2D::StrokePath(def) => c.stroke_path(&path_of(&def)),
                Op2D::ClipPath(def, rule) => c.clip_path(
                    &path_of(&def),
                    FillRule::parse(&rule).unwrap_or(FillRule::NonZero),
                ),
                Op2D::AppendPath(def) => c.append_path(&path_of(&def)),
            }
        });
    }
}

/// Install htmlbox as the surface `web:canvas` paints into.
pub fn install() {
    canvas_backend::set_backend(Arc::new(HtmlBoxBackend));
}

/// The shadow currently in effect, rebuilt from the four attributes.
///
/// The seam sets shadow components one at a time because the IDL has four
/// separate attributes; the engine holds them as one value. Without this,
/// assigning `shadowBlur` would silently reset `shadowColor` to its default.
fn current_shadow(c: &mut dyn rhtmledit::canvas::Canvas) -> Shadow {
    Shadow {
        color: parse_serialized_color(&c.shadow_color_css()),
        blur: c.shadow_blur(),
        offset_x: c.shadow_offset_x(),
        offset_y: c.shadow_offset_y(),
    }
}

/// A serialized colour back to a colour.
///
/// Only ever sees what `shadow_color_css` produced — `#rrggbb` or
/// `rgba(r, g, b, a)` — so it parses those two and nothing else.
fn parse_serialized_color(css: &str) -> Color {
    let css = css.trim();
    if let Some(hex) = css.strip_prefix('#') {
        if hex.len() == 6 {
            if let Ok(v) = u32::from_str_radix(hex, 16) {
                return Color {
                    r: (v >> 16) as u8,
                    g: (v >> 8) as u8,
                    b: v as u8,
                    a: 255,
                };
            }
        }
    }
    if let Some(inner) = css.strip_prefix("rgba(").and_then(|s| s.strip_suffix(')')) {
        let parts: Vec<&str> = inner.split(',').map(str::trim).collect();
        if parts.len() == 4 {
            let n = |i: usize| parts[i].parse::<f32>().unwrap_or(0.0);
            return Color {
                r: n(0) as u8,
                g: n(1) as u8,
                b: n(2) as u8,
                a: (n(3) * 255.0).round().clamp(0.0, 255.0) as u8,
            };
        }
    }
    // Transparent black is the initial `shadowColor`, so an unreadable value
    // lands on the value a canvas starts with rather than on opaque black.
    Color { r: 0, g: 0, b: 0, a: 0 }
}

/// The seam's gradient definition, in the engine's terms.
fn gradient(def: &GradientDef) -> CanvasGradient {
    let mut g = match def.kind {
        SeamGradientKind::Linear { x0, y0, x1, y1 } => CanvasGradient::linear(x0, y0, x1, y1),
        SeamGradientKind::Radial { x0, y0, r0, x1, y1, r1 } => {
            CanvasGradient::radial(x0, y0, r0, x1, y1, r1)
        }
        SeamGradientKind::Conic { angle, x, y } => CanvasGradient::conic(angle, x, y),
    };
    for (offset, css) in &def.stops {
        // A stop whose colour will not parse is DROPPED: `addColorStop` throws
        // on one, so a gradient can never contain it.
        if let Some(color) = parse_stop_color(css) {
            g.stops.push(ColorStop { offset: *offset, color });
        }
    }
    g
}

/// The seam's pattern definition, in the engine's terms.
fn pattern(def: PatternDef) -> CanvasPattern {
    CanvasPattern {
        image: Image::from_rgba(def.width, def.height, def.pixels),
        repetition: Repetition::parse(&def.repetition).unwrap_or(Repetition::Repeat),
    }
}

/// A CSS `<color>` for a gradient stop, through the engine's parser.
fn parse_stop_color(css: &str) -> Option<rhtmledit::canvas::Color> {
    rhtmledit::canvas::parse_color_css(css)
}

/// The seam's path definition, in the engine's terms.
///
/// One arm per segment kind, so a segment the seam learns to carry and this
/// forgets to convert is a compile error — not a shape that quietly does not
/// draw.
fn path_of(def: &PathDef) -> CanvasPath {
    let mut path = CanvasPath::default();
    for op in &def.ops {
        path.ops.push(match *op {
            PathOp2D::ClosePath => EnginePathOp::ClosePath,
            PathOp2D::MoveTo(x, y) => EnginePathOp::MoveTo(x, y),
            PathOp2D::LineTo(x, y) => EnginePathOp::LineTo(x, y),
            PathOp2D::QuadraticCurveTo { cx, cy, x, y } => {
                EnginePathOp::QuadraticCurveTo { cx, cy, x, y }
            }
            PathOp2D::BezierCurveTo { cx1, cy1, cx2, cy2, x, y } => {
                EnginePathOp::BezierCurveTo { cx1, cy1, cx2, cy2, x, y }
            }
            PathOp2D::ArcTo { x1, y1, x2, y2, radius } => {
                EnginePathOp::ArcTo { x1, y1, x2, y2, radius }
            }
            PathOp2D::Rect { x, y, w, h } => EnginePathOp::Rect { x, y, w, h },
            PathOp2D::RoundRect { x, y, w, h, radii } => {
                EnginePathOp::RoundRect { x, y, w, h, radii }
            }
            PathOp2D::Arc { x, y, r, start, end, ccw } => {
                EnginePathOp::Arc { x, y, r, start, end, ccw }
            }
            PathOp2D::Ellipse { x, y, rx, ry, rotation, start, end, ccw } => {
                EnginePathOp::Ellipse { x, y, rx, ry, rotation, start, end, ccw }
            }
        });
    }
    path
}
