//! LinkLabel widget — label with underline, like a hyperlink.

use tiny_skia::*;
use super::WidgetColors;

pub struct LinkLabel {
    pub text: String,
    pub width: f32,
    pub height: f32,
    pub hovered: bool,
    pub visited: bool,
    pub colors: WidgetColors,
}

impl LinkLabel {
    pub fn new<S: Into<String>>(text: S) -> Self {
        Self {
            text: text.into(),
            width: 100.0,
            height: 20.0,
            hovered: false,
            visited: false,
            colors: WidgetColors {
                foreground: (0, 102, 204, 255),
                background: (0, 0, 0, 0), // transparent
                ..WidgetColors::default()
            },
        }
    }

    /// Paint the link label — underline beneath text area. Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Underline path beneath text area
        let underline_y = y + self.height - 2.0;
        let link_color = if self.visited {
            (128, 0, 128, 255) // purple for visited
        } else if self.hovered {
            (0, 70, 180, 255) // darker blue on hover
        } else {
            self.colors.foreground
        };
        let (r, g, b, a) = link_color;
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        let mut pb = PathBuilder::new();
        pb.move_to(x, underline_y);
        pb.line_to(x + self.width, underline_y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Handle click — returns true if within bounds.
    pub fn click(&self, x: f32, y: f32) -> bool {
        x >= 0.0 && y >= 0.0 && x <= self.width && y <= self.height
    }
}
