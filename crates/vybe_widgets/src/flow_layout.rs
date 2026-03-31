//! FlowLayoutPanel widget — container with dashed border (designer preview).

use tiny_skia::*;
use super::WidgetColors;

pub struct FlowLayoutPanel {
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl FlowLayoutPanel {
    pub fn new() -> Self {
        Self {
            width: 300.0,
            height: 200.0,
            colors: WidgetColors {
                background: (250, 250, 250, 255),
                border: (180, 180, 180, 255),
                ..WidgetColors::default()
            },
        }
    }

    /// Paint — light background with dashed border (designer preview style).
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Light background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Dashed border
        let (r, g, b, a) = self.colors.border;
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        stroke.dash = StrokeDash::new(vec![4.0, 3.0], 0.0);

        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Flow direction arrow hint (small right-pointing arrow in top-left)
        paint.set_color_rgba8(r, g, b, 120);
        let mut solid_stroke = Stroke::default();
        solid_stroke.width = 1.0;
        let ax = x + 8.0;
        let ay = y + 8.0;
        let mut pb = PathBuilder::new();
        pb.move_to(ax, ay);
        pb.line_to(ax + 10.0, ay);
        pb.move_to(ax + 7.0, ay - 3.0);
        pb.line_to(ax + 10.0, ay);
        pb.line_to(ax + 7.0, ay + 3.0);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &solid_stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}
