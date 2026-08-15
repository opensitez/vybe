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
