//! Progress bar widget

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};

pub struct ProgressBar {
    pub value: f32, // 0.0..1.0
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl ProgressBar {
    pub fn new() -> Self {
        Self { value: 0.0, width: 200.0, height: 16.0, colors: WidgetColors::default() }
    }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = (240, 240, 240, 255);
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 4.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Filled portion
        let fill_w = (self.value.clamp(0.0, 1.0)) * self.width;
        let (fr, fg, fb, fa) = self.colors.accent;
        paint.set_color_rgba8(fr, fg, fb, fa);
        if fill_w > 0.0 {
            if let Some(path) = rounded_rect_path(x, y, fill_w, self.height, 4.0) {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }
        }

        // Border
        let (br, bg, bb, ba) = self.colors.border;
        paint.set_color_rgba8(br, bg, bb, ba);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 4.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) { (self.width, self.height) }
}
