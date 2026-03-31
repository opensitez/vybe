//! StatusStrip widget — colored status bar at bottom of window.

use tiny_skia::*;
use super::WidgetColors;

pub struct StatusStrip {
    pub text: String,
    pub items: Vec<String>,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl StatusStrip {
    pub fn new() -> Self {
        Self {
            text: String::new(),
            items: Vec::new(),
            width: 800.0,
            height: 24.0,
            colors: WidgetColors {
                background: (0, 122, 204, 255), // VS Code blue
                foreground: (255, 255, 255, 255),
                ..WidgetColors::default()
            },
        }
    }

    /// Paint — colored bar with top border. Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Top border (slightly darker)
        let dr = r.saturating_sub(20);
        let dg = g.saturating_sub(20);
        let db = b.saturating_sub(20);
        paint.set_color_rgba8(dr, dg, db, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + self.width, y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Item separator lines (if multiple items)
        if self.items.len() > 1 {
            // Lighter separator
            paint.set_color_rgba8(r.min(235) + 20, g.min(235) + 20, b.min(235) + 20, 100);
            let item_w = self.width / self.items.len() as f32;
            for i in 1..self.items.len() {
                let sx = x + i as f32 * item_w;
                let mut pb = PathBuilder::new();
                pb.move_to(sx, y + 4.0);
                pb.line_to(sx, y + self.height - 4.0);
                if let Some(path) = pb.finish() {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}
