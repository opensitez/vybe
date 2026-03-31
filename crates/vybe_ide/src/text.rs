//! Shared text rendering helpers using cosmic-text.
//!
//! All public functions take **logical** coordinates and a scale factor.
//! They convert to physical pixels internally.

use cosmic_text::{
    Attrs, Buffer, Color as CosmicColor, Family, FontSystem, Metrics, SwashCache,
};
use tiny_skia::{Pixmap, PixmapPaint, Transform};

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
    let metrics = Metrics::new(font_size, font_size * 1.3).scale(scale);
    let mut buf = Buffer::new(fs, metrics);
    buf.set_text(
        fs,
        text,
        &Attrs::new().family(Family::SansSerif).color(color),
        cosmic_text::Shaping::Advanced,
        None,
    );
    buf.shape_until_scroll(fs, false);
    draw_buffer(pix, fs, sc, &buf, x * scale, y * scale, color);
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
    let metrics = Metrics::new(font_size, font_size * 1.3).scale(scale);
    let mut buf = Buffer::new(fs, metrics);
    buf.set_text(
        fs,
        text,
        &Attrs::new().family(Family::Monospace).color(color),
        cosmic_text::Shaping::Advanced,
        None,
    );
    buf.shape_until_scroll(fs, false);
    draw_buffer(pix, fs, sc, &buf, x * scale, y * scale, color);
}

/// Draw a pre-shaped cosmic-text buffer at physical pixel position (px, py).
fn draw_buffer(
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
