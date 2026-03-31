//! Simple data grid widget — header + rows grid rendering.

use tiny_skia::*;
use super::{WidgetColors};

pub struct DataGrid {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<String>>,
    pub width: f32,
    pub height: f32,
    pub row_height: f32,
    pub header_height: f32,
    pub colors: WidgetColors,
}

impl DataGrid {
    pub fn new(cols: &[&str]) -> Self {
        Self {
            columns: cols.iter().map(|s| s.to_string()).collect(),
            rows: Vec::new(),
            width: 400.0,
            height: 200.0,
            row_height: 20.0,
            header_height: 24.0,
            colors: WidgetColors::default(),
        }
    }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        paint.set_color_rgba8(255, 255, 255, 255);
        let rect = Rect::from_xywh(x, y, self.width, self.height).unwrap();
        pixmap.fill_rect(rect, &paint, ts, None);

        // Header background
        paint.set_color_rgba8(245, 245, 245, 255);
        let header_rect = Rect::from_xywh(x, y, self.width, self.header_height).unwrap();
        pixmap.fill_rect(header_rect, &paint, ts, None);

        // Vertical column separators and header text omitted (text handled by external font engine)
        let col_w = if self.columns.is_empty() { self.width } else { self.width / self.columns.len() as f32 };
        // Draw vertical lines
        paint.set_color_rgba8(200, 200, 200, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        for i in 0..=self.columns.len() {
            let cx = x + i as f32 * col_w;
            let mut pb = PathBuilder::new();
            pb.move_to(cx, y);
            pb.line_to(cx, y + self.height);
            if let Some(path) = pb.finish() { pixmap.stroke_path(&path, &paint, &stroke, ts, None); }
        }

        // Horizontal lines for header + rows
        let mut row_y = y + self.header_height;
        let mut line_count = 0;
        while row_y <= y + self.height + 0.01 {
            let mut pb = PathBuilder::new();
            pb.move_to(x, row_y);
            pb.line_to(x + self.width, row_y);
            if let Some(path) = pb.finish() { pixmap.stroke_path(&path, &paint, &stroke, ts, None); }
            row_y += self.row_height;
            line_count += 1;
            if line_count > 1000 { break; }
        }
    }

    pub fn measure(&self) -> (f32, f32) { (self.width, self.height) }
}
