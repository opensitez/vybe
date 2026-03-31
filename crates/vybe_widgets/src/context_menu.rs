//! ContextMenu widget — popup menu with shadow and hover highlight.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};

pub struct ContextMenu {
    pub items: Vec<String>,
    pub visible: bool,
    pub hover_index: Option<usize>,
    pub item_height: f32,
    pub width: f32,
    pub colors: WidgetColors,
}

impl ContextMenu {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            visible: false,
            hover_index: None,
            item_height: 24.0,
            width: 160.0,
            colors: WidgetColors::default(),
        }
    }

    /// Computed height based on item count.
    pub fn height(&self) -> f32 {
        let count = self.items.len().max(1) as f32;
        count * self.item_height + 4.0 // 2px padding top+bottom
    }

    /// Paint the context menu — white popup with shadow, items with hover highlight.
    /// Item text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        if !self.visible { return; }

        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let h = self.height();

        // Drop shadow (offset by 2px)
        paint.set_color_rgba8(0, 0, 0, 40);
        if let Some(path) = rounded_rect_path(x + 2.0, y + 2.0, self.width, h, 2.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // White background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(path) = rounded_rect_path(x, y, self.width, h, 2.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Hover highlight
        if let Some(idx) = self.hover_index {
            if idx < self.items.len() {
                let item_y = y + 2.0 + idx as f32 * self.item_height;
                let (r, g, b, _) = self.colors.accent;
                paint.set_color_rgba8(r, g, b, 30);
                if let Some(rect) = Rect::from_xywh(x + 2.0, item_y, self.width - 4.0, self.item_height) {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
            }
        }

        // Separator lines between items (light gray)
        paint.set_color_rgba8(230, 230, 230, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        for i in 1..self.items.len() {
            let sep_y = y + 2.0 + i as f32 * self.item_height;
            let mut pb = PathBuilder::new();
            pb.move_to(x + 4.0, sep_y);
            pb.line_to(x + self.width - 4.0, sep_y);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }

        // Border
        paint.set_color_rgba8(180, 180, 180, 255);
        if let Some(path) = rounded_rect_path(x, y, self.width, h, 2.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height())
    }

    /// Hit test — returns item index at (mx, my) relative to menu origin.
    pub fn hit_test(&self, mx: f32, my: f32) -> Option<usize> {
        if !self.visible { return None; }
        if mx < 0.0 || mx > self.width || my < 2.0 || my > self.height() - 2.0 {
            return None;
        }
        let idx = ((my - 2.0) / self.item_height) as usize;
        if idx < self.items.len() { Some(idx) } else { None }
    }

    /// Update hover on mouse move.
    pub fn mouse_move(&mut self, mx: f32, my: f32) {
        self.hover_index = self.hit_test(mx, my);
    }

    /// Y position of item for text placement.
    pub fn item_y(&self, index: usize) -> f32 {
        2.0 + index as f32 * self.item_height
    }
}
