//! Label widget — standalone tiny-skia rendered label.

use tiny_skia::*;
use super::WidgetColors;

pub struct Label {
    pub text: String,
    pub width: f32,
    pub height: f32,
    pub auto_size: bool,
    pub transparent: bool,
    pub colors: WidgetColors,
}

impl Label {
    pub fn new<S: Into<String>>(text: S) -> Self {
        Self {
            text: text.into(),
            width: 100.0,
            height: 20.0,
            auto_size: true,
            transparent: true,
            colors: WidgetColors {
                background: (240, 240, 240, 255),
                ..WidgetColors::default()
            },
        }
    }

    /// Paint the label background (if not transparent). Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        if self.transparent {
            return;
        }
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background fill
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}
