//! ListBox widget — standalone tiny-skia rendered list box.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};

pub struct ListBox {
    pub items: Vec<String>,
    pub selected_index: Option<usize>,
    pub item_height: f32,
    pub scroll_offset: f32,
    pub focused: bool,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl ListBox {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            selected_index: None,
            item_height: 18.0,
            scroll_offset: 0.0,
            focused: false,
            width: 120.0,
            height: 120.0,
            colors: WidgetColors::default(),
        }
    }

    /// Paint the listbox — white background, inset border, selection highlight.
    /// Item text is drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // White background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 1.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Selection highlight bar
        if let Some(idx) = self.selected_index {
            let item_y = y + 1.0 + (idx as f32 * self.item_height) - self.scroll_offset;
            let bar_top = item_y.max(y + 1.0);
            let bar_bottom = (item_y + self.item_height).min(y + self.height - 1.0);
            if bar_top < bar_bottom {
                // Accent blue highlight
                let (r, g, b, _) = self.colors.accent;
                paint.set_color_rgba8(r, g, b, 60);
                if let Some(rect) = Rect::from_xywh(x + 1.0, bar_top, self.width - 2.0, bar_bottom - bar_top) {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
                // Highlight border
                paint.set_color_rgba8(r, g, b, 160);
                let mut stroke = Stroke::default();
                stroke.width = 1.0;
                if let Some(rect_path) = rounded_rect_path(x + 1.0, bar_top, self.width - 2.0, bar_bottom - bar_top, 0.0) {
                    pixmap.stroke_path(&rect_path, &paint, &stroke, ts, None);
                }
            }
        }

        // Inset border (3D sunken effect: dark top-left, light bottom-right)
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        // Dark top and left edges (inset)
        paint.set_color_rgba8(130, 135, 144, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x, y);
        pb.line_to(x + self.width, y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Light bottom and right edges
        paint.set_color_rgba8(255, 255, 255, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Inner border
        paint.set_color_rgba8(160, 160, 160, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x + 1.0, y + self.height - 1.0);
        pb.line_to(x + 1.0, y + 1.0);
        pb.line_to(x + self.width - 1.0, y + 1.0);
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

    /// Handle click at (x, y) relative to widget origin. Returns item index if hit.
    pub fn click(&mut self, x: f32, y: f32) -> Option<usize> {
        if x < 0.0 || y < 0.0 || x > self.width || y > self.height {
            return None;
        }
        let adjusted_y = y - 1.0 + self.scroll_offset;
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

    /// Y position of an item relative to widget origin (for text placement).
    pub fn item_y(&self, index: usize) -> f32 {
        1.0 + (index as f32 * self.item_height) - self.scroll_offset
    }
}
