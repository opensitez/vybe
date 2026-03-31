//! ListView widget — column-based list with header row.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};

pub struct ListView {
    pub items: Vec<String>,
    pub columns: Vec<String>,
    pub selected_index: Option<usize>,
    pub item_height: f32,
    pub header_height: f32,
    pub scroll_offset: f32,
    pub focused: bool,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl ListView {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            columns: Vec::new(),
            selected_index: None,
            item_height: 20.0,
            header_height: 24.0,
            scroll_offset: 0.0,
            focused: false,
            width: 200.0,
            height: 150.0,
            colors: WidgetColors::default(),
        }
    }

    /// Paint the list view — white background, header, column dividers, selection.
    /// Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // White background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Header background (darker gray)
        paint.set_color_rgba8(230, 230, 230, 255);
        if let Some(rect) = Rect::from_xywh(x + 1.0, y + 1.0, self.width - 2.0, self.header_height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Header bottom border
        paint.set_color_rgba8(180, 180, 180, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.header_height);
        pb.line_to(x + self.width, y + self.header_height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Column dividers
        if !self.columns.is_empty() {
            let col_w = self.width / self.columns.len() as f32;
            paint.set_color_rgba8(200, 200, 200, 255);
            for i in 1..self.columns.len() {
                let cx = x + i as f32 * col_w;
                let mut pb = PathBuilder::new();
                pb.move_to(cx, y);
                pb.line_to(cx, y + self.height);
                if let Some(path) = pb.finish() {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
            }
        }

        // Selection highlight
        if let Some(idx) = self.selected_index {
            let item_y = y + self.header_height + 1.0 + (idx as f32 * self.item_height) - self.scroll_offset;
            let bar_top = item_y.max(y + self.header_height + 1.0);
            let bar_bottom = (item_y + self.item_height).min(y + self.height - 1.0);
            if bar_top < bar_bottom {
                let (r, g, b, _) = self.colors.accent;
                paint.set_color_rgba8(r, g, b, 50);
                if let Some(rect) = Rect::from_xywh(x + 1.0, bar_top, self.width - 2.0, bar_bottom - bar_top) {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
            }
        }

        // Inset border (sunken)
        paint.set_color_rgba8(130, 135, 144, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x, y);
        pb.line_to(x + self.width, y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
        paint.set_color_rgba8(255, 255, 255, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Focus ring
        if self.focused {
            let (r, g, b, a) = self.colors.focus_ring;
            paint.set_color_rgba8(r, g, b, a);
            stroke.width = 2.0;
            if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 1.0) {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Handle click — returns item index.
    pub fn click(&mut self, x: f32, y: f32) -> Option<usize> {
        if x < 0.0 || y < 0.0 || x > self.width || y > self.height {
            return None;
        }
        let adjusted_y = y - self.header_height - 1.0 + self.scroll_offset;
        if adjusted_y < 0.0 {
            return None;
        }
        let idx = (adjusted_y / self.item_height) as usize;
        if idx < self.items.len() {
            self.selected_index = Some(idx);
            Some(idx)
        } else {
            None
        }
    }

    /// Column width for layout.
    pub fn column_width(&self) -> f32 {
        if self.columns.is_empty() {
            self.width
        } else {
            self.width / self.columns.len() as f32
        }
    }
}
