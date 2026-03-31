//! Status bar at the bottom of the IDE.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform};

use crate::layout::Rect;
use crate::text::draw_text;

pub struct StatusBar {
    pub message: String,
}

impl StatusBar {
    pub fn new() -> Self {
        Self { message: "Ready".to_string() }
    }

    pub fn render(&self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: Rect, scale: f32) {
        let s = scale;
        let mut paint = Paint::default();

        // Background (blue like VS Code)
        paint.set_color_rgba8(0, 102, 204, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(rect.x * s, rect.y * s, rect.w * s, rect.h * s) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }

        let text_color = CosmicColor::rgba(255, 255, 255, 255);
        draw_text(pix, fs, sc, &self.message, rect.x + 10.0, rect.y + 4.0, 12.0, text_color, s);
    }
}
