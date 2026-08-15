//! Shared text rendering helpers using cosmic-text.
//!
//! All public functions take **logical** coordinates and a scale factor.
//! They convert to physical pixels internally.

use cosmic_text::{
    Attrs, Buffer, Color as CosmicColor, Family, FontSystem, Metrics, Style, SwashCache, Weight,
};
use tiny_skia::{Pixmap, PixmapPaint, Rect, Transform};

/// A resolved text style — everything shaping needs that is not the text itself.
///
/// The family is a **string**, not a `cosmic_text::Family`, because that is
/// what CSS gives us and because a borrowed `Family<'a>` cannot be stored on a
/// widget. Resolution to a font happens at draw time, in [`family_of`].
///
/// This exists so no draw site has to name a font. Every one of them used to:
/// `Family::Monospace` in `layout.rs`, `Family::SansSerif` here — which meant a
/// declared `font-family` had nowhere to arrive, and `font-weight` and
/// `font-style` had no channel at all.
#[derive(Clone, Debug, PartialEq)]
pub struct FontSpec {
    /// `font-family`, first name only. Empty means the generic default below.
    pub family: String,
    /// `font-size`, in logical pixels.
    pub size: f32,
    /// CSS numeric weight — 400 normal, 700 bold.
    pub weight: u16,
    /// `font-style: italic` (or `oblique`, which we render the same).
    pub italic: bool,
    /// `text-decoration: underline`. cosmic-text does not draw decorations, so
    /// these are struck as rules after the glyphs.
    pub underline: bool,
    /// `text-decoration: line-through`.
    pub line_through: bool,
    /// `line-height`, in logical pixels. `None` is CSS `normal`, which every
    /// engine resolves to a font-dependent multiple of the size; 1.3 is what
    /// this toolkit has always used.
    ///
    /// It belongs here rather than at the draw site because the line box is
    /// what a baseline is measured within — two callers using the same size and
    /// different line heights put the same glyphs on different rows.
    pub line_height: Option<f32>,
}

impl FontSpec {
    /// The toolkit's default UI text: whatever the platform calls sans-serif.
    pub fn sans(size: f32) -> Self {
        Self {
            family: String::new(),
            size,
            weight: 400,
            italic: false,
            underline: false,
            line_through: false,
            line_height: None,
        }
    }

    /// The line box height this spec resolves to.
    pub fn resolved_line_height(&self) -> f32 {
        self.line_height.unwrap_or(self.size * 1.3)
    }

    pub fn with_line_height(mut self, line_height: f32) -> Self {
        self.line_height = Some(line_height);
        self
    }

    /// Fixed-pitch text — code editors, the debugger's own chrome.
    pub fn mono(size: f32) -> Self {
        Self {
            family: "monospace".to_string(),
            ..Self::sans(size)
        }
    }

    pub fn with_weight(mut self, weight: u16) -> Self {
        self.weight = weight;
        self
    }

    pub fn with_italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }

    /// Take the whole font axis from a resolved computed style.
    ///
    /// The other direction from [`FontSpec::apply_command`], and needed because
    /// an inline run is produced from a **computed style**, never from a
    /// declaration: the cascade has already run by the time a run exists, so
    /// there is no command to consume. Same mapping, stated once, so a run and
    /// a widget cannot disagree about what `bold` is.
    ///
    /// Anything the cascade did not specify is left as it was, which is what
    /// makes this a merge rather than a replacement.
    pub fn apply_computed(&mut self, props: &crate::css::CssProperties) {
        if let Some(family) = &props.font_family {
            // `font-family` is a LIST and the first name wins — the fallback
            // chain needs a font database to walk.
            let first = family
                .split(',')
                .next()
                .unwrap_or("")
                .trim()
                .trim_matches(|c| c == '"' || c == '\'');
            if !first.is_empty() {
                self.family = first.to_string();
            }
        }
        if let Some(size) = props.font_size {
            self.size = size;
        }
        if let Some(weight) = props.font_weight {
            self.weight = weight as u16;
        }
        if let Some(style) = props.font_style {
            self.italic = style != crate::css::FontStyle::Normal;
        }
        if let Some(underline) = props.underline {
            self.underline = underline;
        }
        if let Some(line_through) = props.line_through {
            self.line_through = line_through;
        }
        if let Some(height) = props.line_height {
            self.line_height = Some(height);
        }
    }

    /// Consume a `SetFont*` / `SetTextDecoration` command, if it is one.
    ///
    /// Returns whether it was handled, so a widget's `Custom` arm can delegate
    /// the whole font axis in one line instead of carrying five copies of these
    /// branches. Every text-bearing control wants exactly this behaviour, and
    /// five widgets each reimplementing "is `bold` 700?" is how they drift.
    pub fn apply_command(&mut self, key: &str, value: &crate::layout::CommandValue) -> bool {
        use crate::layout::{CommandValue, command_number};
        let text = || match value {
            CommandValue::Text(t) => Some(t.trim().to_ascii_lowercase()),
            _ => None,
        };
        match key {
            "SetFontSize" => {
                if let Some(size) = command_number(value) {
                    self.size = size as f32;
                }
            }
            "SetFontFamily" => {
                if let CommandValue::Text(family) = value {
                    self.family = family.clone();
                }
            }
            // A CSS weight is a NUMBER, and `bold`/`normal` are names for 700
            // and 400 — so both spellings are one property, and `font-weight:
            // 600` stays expressible.
            "SetFontWeight" => {
                if let Some(weight) = command_number(value) {
                    self.weight = weight as u16;
                } else if let Some(name) = text() {
                    self.weight = match name.as_str() {
                        "bold" | "bolder" => 700,
                        "lighter" => 300,
                        _ => 400,
                    };
                }
            }
            "SetFontStyle" => {
                if let Some(style) = text() {
                    self.italic = style == "italic" || style == "oblique";
                }
            }
            // `text-decoration` is a LIST — `underline line-through` asks for
            // both — so each line is asked for independently.
            "SetTextDecoration" => {
                if let Some(decoration) = text() {
                    self.underline = decoration.contains("underline");
                    self.line_through = decoration.contains("line-through");
                }
            }
            _ => return false,
        }
        true
    }

    fn attrs<'a>(&'a self, color: Option<CosmicColor>) -> Attrs<'a> {
        let mut attrs = Attrs::new()
            .family(family_of(&self.family))
            .weight(Weight(self.weight))
            .style(if self.italic {
                Style::Italic
            } else {
                Style::Normal
            });
        if let Some(color) = color {
            attrs = attrs.color(color);
        }
        attrs
    }

    fn draws_a_rule(&self) -> bool {
        self.underline || self.line_through
    }
}

impl Default for FontSpec {
    fn default() -> Self {
        Self::sans(13.0)
    }
}

/// A CSS family name to a font.
///
/// The five CSS generic families are keywords, not font names — asking the
/// system for a face literally called "monospace" finds nothing. Everything
/// else is a real family name.
pub fn family_of(name: &str) -> Family<'_> {
    match name.trim().to_ascii_lowercase().as_str() {
        "" | "sans-serif" | "system-ui" | "ui-sans-serif" => Family::SansSerif,
        "serif" | "ui-serif" => Family::Serif,
        "monospace" | "ui-monospace" => Family::Monospace,
        "cursive" => Family::Cursive,
        "fantasy" => Family::Fantasy,
        _ => Family::Name(name.trim()),
    }
}

/// Draw a single line of sans-serif text.
/// `x`, `y` are in **logical** pixels; `font_size` is in logical points.
pub fn draw_text(
    pix: &mut Pixmap,
    fs: &mut FontSystem,
    sc: &mut SwashCache,
    text: &str,
    x: f32,
    y: f32,
    font_size: f32,
    color: CosmicColor,
    scale: f32,
) {
    draw_text_with_font(pix, fs, sc, text, x, y, None, font_size, color, scale);
}

/// Draw a single line of text, parsing an optional font string like "Roboto, 14px"
/// Fallbacks to SansSerif and default_size if none provided.
pub fn draw_text_with_font(
    pix: &mut Pixmap,
    fs: &mut FontSystem,
    sc: &mut SwashCache,
    text: &str,
    x: f32,
    y: f32,
    font_prop: Option<&str>,
    default_size: f32,
    color: CosmicColor,
    scale: f32,
) {
    let (family, size) = parse_font_prop(font_prop, default_size);
    let spec = FontSpec {
        family,
        size,
        ..FontSpec::sans(size)
    };
    draw_text_spec(pix, fs, sc, text, x, y, &spec, color, scale);
}

/// Draw a single line in a fully specified style, at **logical** coordinates.
///
/// This is the one shaping path; everything else in this module is a
/// convenience that builds a [`FontSpec`] and calls it.
pub fn draw_text_spec(
    pix: &mut Pixmap,
    fs: &mut FontSystem,
    sc: &mut SwashCache,
    text: &str,
    x: f32,
    y: f32,
    spec: &FontSpec,
    color: CosmicColor,
    scale: f32,
) {
    draw_text_spec_physical(pix, fs, sc, text, x * scale, y * scale, spec, color, scale);
}

/// As [`draw_text_spec`], but `px`/`py` are already in **physical** pixels
/// while the metrics still scale.
///
/// The two coordinate conventions both exist in the toolkit — `RenderContext`
/// works in physical pixels, this module in logical ones — and converting
/// between them by dividing would round. So the shaping is shared and only the
/// origin differs.
pub fn draw_text_spec_physical(
    pix: &mut Pixmap,
    fs: &mut FontSystem,
    sc: &mut SwashCache,
    text: &str,
    px: f32,
    py: f32,
    spec: &FontSpec,
    color: CosmicColor,
    scale: f32,
) {
    let metrics = Metrics::new(spec.size, spec.resolved_line_height()).scale(scale);
    let mut buf = Buffer::new(fs, metrics);
    buf.set_text(
        fs,
        text,
        &spec.attrs(Some(color)),
        cosmic_text::Shaping::Advanced,
        None,
    );
    buf.shape_until_scroll(fs, false);
    draw_buffer(pix, fs, sc, &buf, px, py, color);
    if spec.draws_a_rule() {
        draw_decorations(pix, &buf, px, py, spec, color, scale);
    }
}

/// Draw differently-styled spans as ONE line of text.
///
/// The whole point of an inline formatting context: `a <strong>b</strong> c`
/// is one line whose middle word is bold, not three stacked boxes. Drawing the
/// spans with three separate `draw_text_spec` calls could not produce it —
/// each call shapes its own buffer from x, so they would overlap at the same
/// origin and none of them would know the others' advances.
///
/// `Buffer::set_rich_text` shapes them together, which is also what makes the
/// line break in the right place: the break opportunities are decided across
/// the whole run sequence rather than per span.
///
/// Returns the width actually laid out, because a caller that centres or
/// right-aligns the line cannot know it any other way.
pub fn draw_rich_text(
    pix: &mut Pixmap,
    fs: &mut FontSystem,
    sc: &mut SwashCache,
    spans: &[(String, FontSpec, CosmicColor)],
    x: f32,
    y: f32,
    wrap_width: Option<f32>,
    scale: f32,
) -> f32 {
    let Some(buf) = shape_rich_text(fs, spans, wrap_width, scale) else {
        return 0.0;
    };
    let width = buf.layout_runs().map(|r| r.line_w).fold(0.0f32, f32::max);
    draw_buffer(
        pix,
        fs,
        sc,
        &buf,
        x * scale,
        y * scale,
        CosmicColor::rgb(0, 0, 0),
    );
    width / scale
}

/// Shape differently-styled spans into one buffer — **the** inline layout.
///
/// Pulled out of [`draw_rich_text`] so that measuring and painting are the same
/// computation and not two of them. A box's height comes from where the lines
/// broke; if the measurement shaped the runs its own way, the height would be
/// right for a layout nobody ever drew.
///
/// `wrap_width` is the content width in LOGICAL pixels, or `None` for a line
/// that never breaks. This is the whole of line breaking: cosmic-text picks the
/// opportunities across the span sequence, which is why the spans have to be
/// shaped together rather than one at a time.
fn shape_rich_text(
    fs: &mut FontSystem,
    spans: &[(String, FontSpec, CosmicColor)],
    wrap_width: Option<f32>,
    scale: f32,
) -> Option<Buffer> {
    let (_, first, _) = spans.first()?;
    // The line box takes the TALLEST span's metrics — a bold 20px word in a
    // 14px sentence sets the line height for the whole line.
    let size = spans.iter().map(|(_, s, _)| s.size).fold(0.0, f32::max);
    let line = spans
        .iter()
        .map(|(_, s, _)| s.resolved_line_height())
        .fold(0.0, f32::max);
    let metrics = Metrics::new(size, line).scale(scale);
    let mut buf = Buffer::new(fs, metrics);
    // The buffer works in physical pixels, as the metrics above do. Height is
    // left unbounded: a block box is as tall as its content, so nothing is
    // clipped away before it has been counted.
    if let Some(w) = wrap_width {
        buf.set_size(fs, Some(w * scale), None);
    }
    let attrs: Vec<(&str, cosmic_text::Attrs)> = spans
        .iter()
        .map(|(text, spec, color)| (text.as_str(), spec.attrs(Some(*color))))
        .collect();
    buf.set_rich_text(
        fs,
        attrs.iter().map(|(t, a)| (*t, a.clone())),
        &first.attrs(None),
        cosmic_text::Shaping::Advanced,
        None,
    );
    buf.shape_until_scroll(fs, false);
    Some(buf)
}

/// The size an inline formatting context takes, in LOGICAL pixels.
///
/// **Intrinsic sizing's measuring half** — CSS §10.6.3: a block box with
/// `height: auto` whose children are inline is as tall as its line boxes, and
/// how many of those there are depends on where the text wrapped. Answers the
/// laid-out width too, which is what a shrink-to-fit box would need; nothing
/// asks for it yet.
///
/// Shaped through [`with_font_system`] at scale 1, because layout is in logical
/// pixels and a box does not change size when the window moves to a denser
/// display.
pub fn measure_rich_text(
    spans: &[(String, FontSpec, CosmicColor)],
    wrap_width: Option<f32>,
) -> (f32, f32) {
    with_font_system(|fs| {
        let Some(buf) = shape_rich_text(fs, spans, wrap_width, 1.0) else {
            return (0.0, 0.0);
        };
        let mut width = 0.0f32;
        let mut height = 0.0f32;
        for run in buf.layout_runs() {
            width = width.max(run.line_w);
            // The BOTTOM of the last line box, not a count of lines times a
            // height: runs differ in height when the spans do, and stacking
            // them by their own tops is what the renderer already does.
            height = height.max(run.line_top + run.line_height);
        }
        (width, height)
    })
}

/// `text-decoration` — the lines cosmic-text does not draw.
///
/// Positioned from the font size rather than from font metrics: underline
/// position and thickness are per-face values this shaping path does not
/// surface, and a fraction of the em is what every toolkit falls back to.
fn draw_decorations(
    pix: &mut Pixmap,
    buf: &Buffer,
    px: f32,
    py: f32,
    spec: &FontSpec,
    color: CosmicColor,
    scale: f32,
) {
    let mut paint = tiny_skia::Paint::default();
    paint.set_color_rgba8(color.r(), color.g(), color.b(), color.a());
    paint.anti_alias = false;
    let thickness = (spec.size * scale / 14.0).max(1.0);
    for run in buf.layout_runs() {
        let width = run.line_w;
        if width <= 0.0 {
            continue;
        }
        let baseline = py + run.line_y;
        // Under the baseline for an underline; through the x-height for a
        // strike, which is what `line-through` means.
        let rules = [
            (spec.underline, baseline + thickness * 2.0),
            (spec.line_through, baseline - spec.size * scale * 0.3),
        ];
        for (wanted, y) in rules {
            if !wanted {
                continue;
            }
            if let Some(rect) = Rect::from_xywh(px, y, width, thickness) {
                paint.anti_alias = false;
                pix.fill_rect(rect, &paint, Transform::identity(), None);
            }
        }
    }
}

/// Parse "Family, Size" string
fn parse_font_prop(font_prop: Option<&str>, default_size: f32) -> (String, f32) {
    if let Some(s) = font_prop {
        let mut parts = s.split(',').map(|p| p.trim());
        let fam = parts.next().unwrap_or("").to_string();
        if let Some(size_str) = parts.next() {
            let num: String = size_str.chars().filter(|c| c.is_ascii_digit()).collect();
            if let Ok(sz) = num.parse::<f32>() {
                if sz > 0.0 {
                    return (fam, sz);
                }
            }
        }
        (fam, default_size)
    } else {
        (String::new(), default_size)
    }
}

/// Measure the logical width of a single line of sans-serif text.
pub fn measure_text(fs: &mut FontSystem, text: &str, font_size: f32, scale: f32) -> f32 {
    measure_text_with_font(fs, text, None, font_size, scale)
}

/// Measure logical width of text, parsing an optional font string.
pub fn measure_text_with_font(
    fs: &mut FontSystem,
    text: &str,
    font_prop: Option<&str>,
    default_size: f32,
    scale: f32,
) -> f32 {
    let (family, size) = parse_font_prop(font_prop, default_size);
    let spec = FontSpec {
        family,
        size,
        ..FontSpec::sans(size)
    };
    measure_text_spec(fs, text, &spec, scale)
}

/// The font database, borrowed for a measurement that has no render pass.
///
/// **Measuring is not painting.** Every `measure_text*` above takes a
/// `&mut FontSystem` because its callers are widgets already inside `render`,
/// holding the one the renderer owns. `measureText` has no such caller: it is
/// a host call from guest code, answered between frames, and a browser answers
/// it without painting anything — font metrics are available to a page that
/// never draws.
///
/// So the database is process-wide and built once, on first use.
/// `FontSystem::new()` enumerates the system's fonts, which is why it is not
/// built per call; a `Mutex` rather than a thread-local because two agents
/// measuring the same font should not each pay for their own enumeration.
///
/// Widgets keep taking the renderer's — this does not replace it, and nothing
/// inside a render pass should reach for it.
pub fn with_font_system<T>(f: impl FnOnce(&mut FontSystem) -> T) -> T {
    static FONTS: std::sync::OnceLock<std::sync::Mutex<FontSystem>> = std::sync::OnceLock::new();
    let fonts = FONTS.get_or_init(|| std::sync::Mutex::new(FontSystem::new()));
    let mut guard = fonts.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    f(&mut guard)
}

/// The logical width of one line in a fully specified style.
///
/// Measuring has to shape with the SAME attributes it will be drawn with —
/// bold and italic faces have different advances, so measuring in the regular
/// face and drawing in bold is how text overruns the box it was sized for.
pub fn measure_text_spec(fs: &mut FontSystem, text: &str, spec: &FontSpec, scale: f32) -> f32 {
    let metrics = Metrics::new(spec.size, spec.resolved_line_height()).scale(scale);
    let mut buf = Buffer::new(fs, metrics);
    buf.set_text(fs, text, &spec.attrs(None), cosmic_text::Shaping::Advanced, None);
    buf.shape_until_scroll(fs, false);

    let mut max_w = 0.0f32;
    for run in buf.layout_runs() {
        max_w = max_w.max(run.line_w);
    }
    max_w / scale
}

/// Draw a single line of monospace text.
/// `x`, `y` are in **logical** pixels; `font_size` is in logical points.
pub fn draw_mono(
    pix: &mut Pixmap,
    fs: &mut FontSystem,
    sc: &mut SwashCache,
    text: &str,
    x: f32,
    y: f32,
    font_size: f32,
    color: CosmicColor,
    scale: f32,
) {
    draw_text_spec(pix, fs, sc, text, x, y, &FontSpec::mono(font_size), color, scale);
}

/// Draw a pre-shaped cosmic-text buffer at physical pixel position (px, py).
///
/// Used by `crate::canvas::tinyskia::TinySkiaCanvas::fill_text` to share
/// the same glyph rasterisation path the rest of the toolkit's text
/// rendering uses. Public to the crate so the canvas module can call
/// it without re-implementing the swash-cache walk.
pub(crate) fn draw_buffer(
    pix: &mut Pixmap,
    fs: &mut FontSystem,
    sc: &mut SwashCache,
    buf: &Buffer,
    px: f32,
    py: f32,
    color: CosmicColor,
) {
    let (cr, cg, cb, ca) = (color.r(), color.g(), color.b(), color.a());
    for run in buf.layout_runs() {
        for glyph in run.glyphs {
            let pg = glyph.physical((px, py + run.line_y), 1.0);
            if let Some(img) = sc.get_image(fs, pg.cache_key) {
                if img.placement.width == 0 || img.placement.height == 0 {
                    continue;
                }
                if let Some(mut glyph_pix) =
                    Pixmap::new(img.placement.width.max(1), img.placement.height.max(1))
                {
                    for (slot, &alpha) in glyph_pix.pixels_mut().iter_mut().zip(img.data.iter()) {
                        let af = (alpha as f32 / 255.0) * (ca as f32 / 255.0);
                        *slot = tiny_skia::ColorU8::from_rgba(
                            (cr as f32 * af) as u8,
                            (cg as f32 * af) as u8,
                            (cb as f32 * af) as u8,
                            (255.0 * af) as u8,
                        )
                        .premultiply();
                    }
                    pix.draw_pixmap(
                        pg.x + img.placement.left,
                        pg.y - img.placement.top,
                        glyph_pix.as_ref(),
                        &PixmapPaint::default(),
                        Transform::identity(),
                        None,
                    );
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_css_generic_families_are_keywords_not_font_names() {
        // The bug this prevents: passing "monospace" to the font database as a
        // family NAME finds no face, because no installed font is called that.
        // The five generics have to become cosmic-text's own variants.
        assert!(matches!(family_of("monospace"), Family::Monospace));
        assert!(matches!(family_of("Monospace"), Family::Monospace));
        assert!(matches!(family_of("serif"), Family::Serif));
        assert!(matches!(family_of("sans-serif"), Family::SansSerif));
        assert!(matches!(family_of("cursive"), Family::Cursive));
        assert!(matches!(family_of("fantasy"), Family::Fantasy));
        // Empty is "unspecified", which is the UI default rather than an error.
        assert!(matches!(family_of(""), Family::SansSerif));
    }

    #[test]
    fn a_real_family_name_is_passed_through() {
        assert!(matches!(family_of("Courier New"), Family::Name("Courier New")));
    }

    /// A width makes the line break; no width makes it run on.
    ///
    /// **The behaviour change to be aware of, stated where it can be seen.**
    /// Every box that paints caption-plus-inline-content now wraps at its
    /// content edge, and that is the CSS rule — a browser breaks a `<div>`'s
    /// text at its box. It is not the toolkit rule: a VCL `TPanel.Caption` is
    /// one centred line however long it is, so a panel with a caption wider
    /// than its box paints differently than it did. Correct, and visible.
    #[test]
    fn a_width_is_what_makes_a_line_break() {
        let spans = vec![(
            "a caption far too long to fit".repeat(4),
            FontSpec::sans(14.0),
            CosmicColor::rgb(0, 0, 0),
        )];
        let (wide, one_line) = measure_rich_text(&spans, None);
        let (narrow, wrapped) = measure_rich_text(&spans, Some(120.0));

        assert!(wide > 120.0, "unwrapped, it is as wide as it needs to be");
        assert!(narrow <= 120.0, "wrapped, it stays inside the box");
        assert!(
            wrapped > one_line * 2.0,
            "and is taller for it: {one_line} -> {wrapped}"
        );
    }

    #[test]
    fn line_height_normal_is_a_multiple_of_the_size_and_an_explicit_one_wins() {
        // `RenderContext::draw_text` depends on this: it has always used a
        // 20px line at 14px text, and resolving that to 14 * 1.3 would move
        // every baseline in the toolkit.
        assert_eq!(FontSpec::sans(10.0).resolved_line_height(), 13.0);
        assert_eq!(
            FontSpec::mono(14.0).with_line_height(20.0).resolved_line_height(),
            20.0
        );
    }
}
