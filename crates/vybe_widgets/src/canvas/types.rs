//! Canvas value types — `Color`, `LineCap`, `LineJoin`, `Font`, `Image`.
//!
//! Kept in their own module so the `Canvas` trait surface and the impls
//! can each pull only what they need. Every type here is `Clone +
//! Debug + PartialEq` so they can live in `DrawCmd` variants without
//! ceremony.

use std::sync::Arc;

/// 32-bit RGBA colour. Components are 0-255.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
    pub a: u8,
}

impl Color {
    pub const BLACK: Color = Color {
        r: 0,
        g: 0,
        b: 0,
        a: 255,
    };
    pub const WHITE: Color = Color {
        r: 255,
        g: 255,
        b: 255,
        a: 255,
    };
    pub const TRANSPARENT: Color = Color {
        r: 0,
        g: 0,
        b: 0,
        a: 0,
    };

    /// Construct an opaque RGB colour.
    pub const fn rgb(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b, a: 255 }
    }

    /// Construct an RGBA colour.
    pub const fn rgba(r: u8, g: u8, b: u8, a: u8) -> Self {
        Self { r, g, b, a }
    }

    /// Convert to `tiny_skia::Color` for backend impls.
    pub fn to_tiny_skia(self) -> tiny_skia::Color {
        tiny_skia::Color::from_rgba8(self.r, self.g, self.b, self.a)
    }
}

/// One entry of a gradient's colour ramp — `CanvasGradient.addColorStop`.
///
/// `offset` is 0..=1 in gradient space. The spec throws `IndexSizeError`
/// outside that range; `add_color_stop` clamps instead, because a recording
/// that silently drops a stop is harder to see than one that pins it to the
/// end. Stops are kept in insertion order and sorted at paint time — the spec
/// requires stable ordering for equal offsets, which a sort-on-insert loses.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ColorStop {
    pub offset: f32,
    pub color: Color,
}

/// How a pattern repeats — `createPattern`'s second argument.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum Repetition {
    #[default]
    Repeat,
    RepeatX,
    RepeatY,
    NoRepeat,
}

impl Repetition {
    /// Parse the spec's four literals. The empty string is `repeat`, per
    /// §"createPattern" — a null or missing repetition is not an error.
    pub fn parse(value: &str) -> Option<Repetition> {
        match value.trim() {
            "" | "repeat" => Some(Repetition::Repeat),
            "repeat-x" => Some(Repetition::RepeatX),
            "repeat-y" => Some(Repetition::RepeatY),
            "no-repeat" => Some(Repetition::NoRepeat),
            _ => None,
        }
    }
}

/// A gradient's geometry. The colour ramp travels beside it in [`Gradient`].
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum GradientKind {
    /// `createLinearGradient(x0, y0, x1, y1)` — the ramp runs along the line.
    Linear { x0: f32, y0: f32, x1: f32, y1: f32 },
    /// `createRadialGradient(x0, y0, r0, x1, y1, r1)` — the ramp runs between
    /// two circles, which is what makes off-centre highlights expressible.
    Radial {
        x0: f32,
        y0: f32,
        r0: f32,
        x1: f32,
        y1: f32,
        r1: f32,
    },
    /// `createConicGradient(startAngle, x, y)` — the ramp sweeps around a
    /// point. Note the spec's argument order puts the ANGLE first.
    Conic { start_angle: f32, x: f32, y: f32 },
}

/// `CanvasGradient` — geometry plus an ordered colour ramp.
#[derive(Clone, Debug, PartialEq)]
pub struct Gradient {
    pub kind: GradientKind,
    pub stops: Vec<ColorStop>,
}

impl Gradient {
    pub fn linear(x0: f32, y0: f32, x1: f32, y1: f32) -> Self {
        Self {
            kind: GradientKind::Linear { x0, y0, x1, y1 },
            stops: Vec::new(),
        }
    }

    pub fn radial(x0: f32, y0: f32, r0: f32, x1: f32, y1: f32, r1: f32) -> Self {
        Self {
            kind: GradientKind::Radial {
                x0,
                y0,
                r0,
                x1,
                y1,
                r1,
            },
            stops: Vec::new(),
        }
    }

    pub fn conic(start_angle: f32, x: f32, y: f32) -> Self {
        Self {
            kind: GradientKind::Conic { start_angle, x, y },
            stops: Vec::new(),
        }
    }

    /// `CanvasGradient.addColorStop(offset, color)`.
    pub fn add_color_stop(&mut self, offset: f32, color: Color) {
        self.stops.push(ColorStop {
            offset: offset.clamp(0.0, 1.0),
            color,
        });
    }

    /// The ramp in paint order: by offset, insertion order preserved for ties.
    ///
    /// `sort_by` is stable, which is the half the spec actually pins down —
    /// two stops at the same offset paint in the order they were added, and
    /// that is how a hard colour boundary is expressed.
    pub fn sorted_stops(&self) -> Vec<ColorStop> {
        let mut stops = self.stops.clone();
        stops.sort_by(|a, b| a.offset.total_cmp(&b.offset));
        stops
    }
}

/// `CanvasPattern` — an image tiled under the shape being painted.
#[derive(Clone, Debug, PartialEq)]
pub struct Pattern {
    pub image: Image,
    pub repetition: Repetition,
}

/// What `fillStyle` / `strokeStyle` hold.
///
/// The spec types both as `DOMString | CanvasGradient | CanvasPattern`, so a
/// bare `Color` cannot represent the state: it makes every gradient and every
/// pattern unexpressible, which is why `LinearGradientBrush` and `HatchBrush`
/// had nowhere to land and stayed on the widget factory.
///
/// `Color` is still the default and still the common case — `Paint::Color` is
/// what a plain `set_fill_color` produces, and the backends keep their fast
/// path for it.
#[derive(Clone, Debug, PartialEq)]
pub enum Paint {
    Color(Color),
    Gradient(Gradient),
    Pattern(Pattern),
}

impl Default for Paint {
    fn default() -> Self {
        Paint::Color(Color::BLACK)
    }
}

impl Paint {
    /// The single colour this paint is, if it is one.
    ///
    /// Backends that cannot build a shader use this to degrade to a flat fill
    /// rather than to nothing: a gradient's FIRST stop is a far better answer
    /// than transparent, and it keeps a shape visible in a capture.
    pub fn as_flat_color(&self) -> Color {
        match self {
            Paint::Color(c) => *c,
            Paint::Gradient(g) => g
                .sorted_stops()
                .first()
                .map(|s| s.color)
                .unwrap_or(Color::TRANSPARENT),
            Paint::Pattern(_) => Color::TRANSPARENT,
        }
    }
}

impl From<Color> for Paint {
    fn from(color: Color) -> Self {
        Paint::Color(color)
    }
}

/// `shadowColor` + `shadowBlur` + `shadowOffsetX` + `shadowOffsetY`.
///
/// The default is a fully transparent colour with no blur and no offset, which
/// is the spec's "no shadow" state — and why [`Shadow::is_visible`] tests the
/// alpha first. A shadow with a transparent colour draws nothing no matter how
/// large the blur.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Shadow {
    pub color: Color,
    pub blur: f32,
    pub offset_x: f32,
    pub offset_y: f32,
}

impl Default for Shadow {
    fn default() -> Self {
        Self {
            color: Color::TRANSPARENT,
            blur: 0.0,
            offset_x: 0.0,
            offset_y: 0.0,
        }
    }
}

impl Shadow {
    /// Does this state paint anything?
    ///
    /// Per the spec's shadow-drawing preconditions: a shadow is painted when
    /// the colour is not fully transparent AND at least one of blur, offsetX
    /// or offsetY is non-zero. A zero-blur zero-offset shadow sits exactly
    /// under the shape and is invisible, so skipping it is not an
    /// approximation — it is the same output for less work.
    pub fn is_visible(&self) -> bool {
        self.color.a != 0 && (self.blur != 0.0 || self.offset_x != 0.0 || self.offset_y != 0.0)
    }
}

/// Which points count as inside a path — `fill(fillRule)` / `clip(fillRule)`.
///
/// This is not cosmetic: a path with self-intersections or a hole punched by a
/// reversed subpath fills DIFFERENTLY under the two rules, and `evenodd` is the
/// only way to express "hole" without splitting the path.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum FillRule {
    #[default]
    NonZero,
    EvenOdd,
}

impl FillRule {
    pub fn parse(value: &str) -> Option<FillRule> {
        match value.trim() {
            "nonzero" => Some(FillRule::NonZero),
            "evenodd" => Some(FillRule::EvenOdd),
            _ => None,
        }
    }

    pub fn to_tiny_skia(self) -> tiny_skia::FillRule {
        match self {
            FillRule::NonZero => tiny_skia::FillRule::Winding,
            FillRule::EvenOdd => tiny_skia::FillRule::EvenOdd,
        }
    }
}

/// How the ends of stroked lines are drawn.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LineCap {
    /// Flat edge ending exactly at the line endpoint.
    Butt,
    /// Half-circle extending past the endpoint.
    Round,
    /// Square extending past the endpoint by half the line width.
    Square,
}

impl Default for LineCap {
    fn default() -> Self {
        LineCap::Butt
    }
}

/// How stroked lines join at corners.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LineJoin {
    /// Sharp corner. Long miters get clipped to `miter_limit`.
    Miter,
    /// Rounded corner.
    Round,
    /// Beveled corner.
    Bevel,
}

impl Default for LineJoin {
    fn default() -> Self {
        LineJoin::Miter
    }
}

/// Font weight (normal vs bold). HTML5 canvas uses numeric weights too,
/// but every backend we care about reduces them to a binary normal/bold
/// distinction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FontWeight {
    Normal,
    Bold,
}

impl Default for FontWeight {
    fn default() -> Self {
        FontWeight::Normal
    }
}

/// Font style (upright vs italic).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FontStyle {
    Normal,
    Italic,
}

/// `textAlign` — which end of the text the `x` in `fillText(text, x, y)` names.
///
/// `start`/`end` are the logical spellings and equal `left`/`right` in a
/// left-to-right context, which is the only direction this canvas lays out.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TextAlign {
    #[default]
    Start,
    End,
    Left,
    Right,
    Center,
}

impl TextAlign {
    pub fn parse(value: &str) -> Option<TextAlign> {
        match value.trim().to_ascii_lowercase().as_str() {
            "start" => Some(TextAlign::Start),
            "end" => Some(TextAlign::End),
            "left" => Some(TextAlign::Left),
            "right" => Some(TextAlign::Right),
            "center" | "centre" => Some(TextAlign::Center),
            _ => None,
        }
    }
}

/// `textBaseline` — which horizontal line of the text the `y` names.
///
/// **The default is `alphabetic`**, and that is not a detail: it means `y` is
/// the baseline, so the glyphs sit ABOVE it. Treating `y` as the top of the
/// text — which this canvas did, with a comment admitting it — puts every
/// string roughly one ascent too low.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TextBaseline {
    Top,
    Hanging,
    Middle,
    #[default]
    Alphabetic,
    Ideographic,
    Bottom,
}

impl TextBaseline {
    pub fn parse(value: &str) -> Option<TextBaseline> {
        match value.trim().to_ascii_lowercase().as_str() {
            "top" => Some(TextBaseline::Top),
            "hanging" => Some(TextBaseline::Hanging),
            "middle" => Some(TextBaseline::Middle),
            "alphabetic" => Some(TextBaseline::Alphabetic),
            "ideographic" => Some(TextBaseline::Ideographic),
            "bottom" => Some(TextBaseline::Bottom),
            _ => None,
        }
    }
}

impl Default for FontStyle {
    fn default() -> Self {
        FontStyle::Normal
    }
}

/// A font specification — family name + size + weight/style.
#[derive(Clone, Debug, PartialEq)]
pub struct Font {
    pub family: String,
    pub size: f32,
    pub weight: FontWeight,
    pub style: FontStyle,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            family: "sans-serif".to_string(),
            size: 12.0,
            weight: FontWeight::Normal,
            style: FontStyle::Normal,
        }
    }
}

impl Font {
    /// Construct a font with the given family and size.
    pub fn new(family: impl Into<String>, size: f32) -> Self {
        Self {
            family: family.into(),
            size,
            weight: FontWeight::Normal,
            style: FontStyle::Normal,
        }
    }

    pub fn with_bold(mut self, bold: bool) -> Self {
        self.weight = if bold {
            FontWeight::Bold
        } else {
            FontWeight::Normal
        };
        self
    }

    pub fn with_italic(mut self, italic: bool) -> Self {
        self.style = if italic {
            FontStyle::Italic
        } else {
            FontStyle::Normal
        };
        self
    }
}

/// Pixel-buffer image, used by `draw_image`.
///
/// Pixels are tightly packed RGBA, row-major. Stored in an `Arc` so
/// `DrawCmd::DrawImage` is cheap to clone (the recording stores image
/// references, not copies of pixel data).
#[derive(Clone, Debug)]
pub struct Image {
    pub width: u32,
    pub height: u32,
    pub pixels: Arc<Vec<u8>>,
}

impl PartialEq for Image {
    fn eq(&self, other: &Self) -> bool {
        self.width == other.width
            && self.height == other.height
            && Arc::ptr_eq(&self.pixels, &other.pixels)
    }
}

impl Image {
    /// The sub-rectangle `(sx, sy, sw, sh)` as a new image — `drawImage`'s
    /// nine-argument source rect.
    ///
    /// Clamped to the image bounds, so a rectangle that hangs off the edge
    /// yields the part that overlaps rather than reading out of bounds. An
    /// empty or fully outside rectangle answers `None`, which the caller
    /// renders as "draw nothing" — the spec's own handling.
    ///
    /// Negative width or height is normalised, matching how the spec reads a
    /// rectangle with either extent negative.
    pub fn crop(&self, sx: f32, sy: f32, sw: f32, sh: f32) -> Option<Image> {
        let (x0, x1) = if sw < 0.0 { (sx + sw, sx) } else { (sx, sx + sw) };
        let (y0, y1) = if sh < 0.0 { (sy + sh, sy) } else { (sy, sy + sh) };

        let x0 = x0.floor().max(0.0) as u32;
        let y0 = y0.floor().max(0.0) as u32;
        let x1 = (x1.ceil().max(0.0) as u32).min(self.width);
        let y1 = (y1.ceil().max(0.0) as u32).min(self.height);
        if x1 <= x0 || y1 <= y0 {
            return None;
        }

        let (w, h) = (x1 - x0, y1 - y0);
        let mut out = Vec::with_capacity((w * h * 4) as usize);
        for row in y0..y1 {
            let start = ((row * self.width + x0) * 4) as usize;
            let end = start + (w * 4) as usize;
            out.extend_from_slice(&self.pixels[start..end]);
        }
        Some(Image::from_rgba(w, h, out))
    }

    /// Construct an image from raw RGBA pixel data.
    pub fn from_rgba(width: u32, height: u32, pixels: Vec<u8>) -> Self {
        debug_assert_eq!(
            pixels.len(),
            (width * height * 4) as usize,
            "Image::from_rgba: expected width*height*4 bytes",
        );
        Self {
            width,
            height,
            pixels: Arc::new(pixels),
        }
    }

    /// Expand an 8-bit palette-indexed image to RGBA.
    ///
    /// This is the shape a software renderer produces: `indices` holds one
    /// byte per pixel, `palette` is up to 256 entries of `0xRRGGBB` (alpha is
    /// forced opaque — an indexed frame has no alpha channel). Doom's
    /// `I_FinishUpdate` is exactly this: an 8-bit screen buffer plus a 256
    /// colour palette, expanded once per frame.
    ///
    /// Lives here rather than in the host bridge because it is an image-format
    /// concern, and it runs NATIVELY — the guest writes indices, the host does
    /// the per-pixel work, so no part of this crosses into WASM.
    ///
    /// Short `indices` are padded with index 0 and out-of-range indices read as
    /// black, so a guest that miscounts gets a visibly wrong frame rather than
    /// a panic.
    pub fn from_paletted(width: u32, height: u32, indices: &[u8], palette: &[u32]) -> Self {
        let count = (width as usize) * (height as usize);
        let mut rgba = Vec::with_capacity(count * 4);
        for i in 0..count {
            let entry = indices
                .get(i)
                .and_then(|&idx| palette.get(idx as usize))
                .copied()
                .unwrap_or(0);
            rgba.push((entry >> 16) as u8);
            rgba.push((entry >> 8) as u8);
            rgba.push(entry as u8);
            rgba.push(0xFF);
        }
        Self {
            width,
            height,
            pixels: Arc::new(rgba),
        }
    }

    /// Decode an image from a file path. Supports PNG, JPEG, GIF, and BMP.
    pub fn from_file(path: impl AsRef<std::path::Path>) -> Result<Self, String> {
        let img = image::open(path.as_ref())
            .map_err(|e| format!("Image::from_file: {}", e))?
            .to_rgba8();
        let (width, height) = img.dimensions();
        Ok(Self::from_rgba(width, height, img.into_raw()))
    }

    /// Decode an image from in-memory bytes (auto-detects format from
    /// content). Supports PNG, JPEG, GIF, and BMP.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, String> {
        let img = image::load_from_memory(bytes)
            .map_err(|e| format!("Image::from_bytes: {}", e))?
            .to_rgba8();
        let (width, height) = img.dimensions();
        Ok(Self::from_rgba(width, height, img.into_raw()))
    }

    /// Save the image as a PNG file.
    pub fn save_png(&self, path: impl AsRef<std::path::Path>) -> Result<(), String> {
        image::save_buffer(
            path.as_ref(),
            &self.pixels,
            self.width,
            self.height,
            image::ColorType::Rgba8,
        )
        .map_err(|e| format!("Image::save_png: {}", e))
    }
}
