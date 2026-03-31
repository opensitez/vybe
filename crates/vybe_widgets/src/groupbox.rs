//! GroupBox widget — border with title gap at top.

use tiny_skia::*;
use super::WidgetColors;

pub struct GroupBox {
    pub title: String,
    /// Estimated width of title text in pixels (caller should set this after measuring text).
    pub title_width: f32,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl GroupBox {
    pub fn new<S: Into<String>>(title: S) -> Self {
        Self {
            title: title.into(),
            title_width: 60.0,
            width: 200.0,
            height: 120.0,
            colors: WidgetColors {
                border: (160, 160, 160, 255),
                ..WidgetColors::default()
            },
        }
    }

    /// Paint the groupbox — border with gap at top for title. Title text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        // The border starts slightly below top to leave room for title text center
        let border_top = y + 8.0;
        let title_gap_start = x + 8.0;
        let title_gap_end = x + 8.0 + self.title_width + 8.0;
        let (r, g, b, a) = self.colors.border;
        paint.set_color_rgba8(r, g, b, a);

        // Top border with gap for title
        // Left part of top border
        let mut pb = PathBuilder::new();
        pb.move_to(x, border_top);
        pb.line_to(title_gap_start, border_top);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Right part of top border (after title gap)
        let mut pb = PathBuilder::new();
        pb.move_to(title_gap_end, border_top);
        pb.line_to(x + self.width, border_top);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Right border
        let mut pb = PathBuilder::new();
        pb.move_to(x + self.width, border_top);
        pb.line_to(x + self.width, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Bottom border
        let mut pb = PathBuilder::new();
        pb.move_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Left border
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x, border_top);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Content area offset (children should be placed inside the border).
    pub fn content_rect(&self) -> (f32, f32, f32, f32) {
        (4.0, 18.0, self.width - 8.0, self.height - 22.0)
    }
}
