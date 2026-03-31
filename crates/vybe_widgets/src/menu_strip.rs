//! MenuStrip widget — horizontal menu bar with item areas.

use tiny_skia::*;
use super::WidgetColors;

pub struct MenuStrip {
    pub items: Vec<String>,
    /// Estimated widths for each item (caller sets from text measurement).
    pub item_widths: Vec<f32>,
    pub active_index: Option<usize>,
    pub hover_index: Option<usize>,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl MenuStrip {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            item_widths: Vec::new(),
            active_index: None,
            hover_index: None,
            width: 400.0,
            height: 24.0,
            colors: WidgetColors {
                background: (240, 240, 240, 255),
                ..WidgetColors::default()
            },
        }
    }

    /// Default item width when item_widths is not set.
    fn default_item_width(&self) -> f32 {
        60.0
    }

    /// Get width for a specific item.
    fn get_item_width(&self, index: usize) -> f32 {
        if index < self.item_widths.len() {
            self.item_widths[index]
        } else {
            self.default_item_width()
        }
    }

    /// X position for a specific item.
    pub fn item_x(&self, index: usize) -> f32 {
        let mut x = 0.0;
        for i in 0..index {
            x += self.get_item_width(i);
        }
        x
    }

    /// Paint — menu bar background, item hover/active highlight, bottom border.
    /// Item text drawn by caller.
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

        // Item highlights
        for i in 0..self.items.len() {
            let ix = x + self.item_x(i);
            let iw = self.get_item_width(i);
            let is_active = self.active_index == Some(i);
            let is_hovered = self.hover_index == Some(i);

            if is_active {
                // Active item: accent background
                let (r, g, b, _) = self.colors.accent;
                paint.set_color_rgba8(r, g, b, 40);
                if let Some(rect) = Rect::from_xywh(ix, y, iw, self.height) {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
            } else if is_hovered {
                // Hover: subtle highlight
                paint.set_color_rgba8(0, 0, 0, 15);
                if let Some(rect) = Rect::from_xywh(ix, y, iw, self.height) {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
            }
        }

        // Bottom border
        paint.set_color_rgba8(210, 210, 210, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x + self.width, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Hit test — returns item index at position.
    pub fn hit_test(&self, mx: f32, _my: f32) -> Option<usize> {
        if _my < 0.0 || _my > self.height { return None; }
        let mut cx = 0.0;
        for i in 0..self.items.len() {
            let iw = self.get_item_width(i);
            if mx >= cx && mx < cx + iw {
                return Some(i);
            }
            cx += iw;
        }
        None
    }

    /// Update hover state on mouse move.
    pub fn mouse_move(&mut self, mx: f32, my: f32) {
        self.hover_index = self.hit_test(mx, my);
    }
}
