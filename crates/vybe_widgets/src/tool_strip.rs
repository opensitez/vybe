//! ToolStrip widget — toolbar with button-like items and separators.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};

#[derive(Clone, Debug)]
pub enum ToolStripItem {
    Button(String),
    Separator,
}

pub struct ToolStrip {
    pub items: Vec<ToolStripItem>,
    pub hover_index: Option<usize>,
    pub pressed_index: Option<usize>,
    pub item_size: f32,
    pub separator_width: f32,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl ToolStrip {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            hover_index: None,
            pressed_index: None,
            item_size: 24.0,
            separator_width: 6.0,
            width: 400.0,
            height: 28.0,
            colors: WidgetColors {
                background: (245, 246, 247, 255),
                ..WidgetColors::default()
            },
        }
    }

    /// X position of each item (returns (x, width) pairs).
    pub fn item_positions(&self) -> Vec<(f32, f32)> {
        let mut positions = Vec::new();
        let mut cx = 2.0;
        for item in &self.items {
            match item {
                ToolStripItem::Button(_) => {
                    positions.push((cx, self.item_size));
                    cx += self.item_size + 1.0;
                }
                ToolStripItem::Separator => {
                    positions.push((cx, self.separator_width));
                    cx += self.separator_width;
                }
            }
        }
        positions
    }

    /// Paint — toolbar background, button items with hover/press, separators.
    /// Button label text/icons drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        // Toolbar background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Bottom border
        paint.set_color_rgba8(210, 210, 210, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x + self.width, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        let positions = self.item_positions();
        let btn_h = self.height - 4.0;
        let btn_y = y + 2.0;

        for (i, item) in self.items.iter().enumerate() {
            if i >= positions.len() { break; }
            let (ix, iw) = positions[i];
            let item_x = x + ix;

            match item {
                ToolStripItem::Button(_) => {
                    let is_pressed = self.pressed_index == Some(i);
                    let is_hovered = self.hover_index == Some(i);

                    if is_pressed {
                        // Pressed: darker background
                        paint.set_color_rgba8(200, 200, 200, 255);
                        if let Some(path) = rounded_rect_path(item_x, btn_y, iw, btn_h, 2.0) {
                            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
                        }
                        paint.set_color_rgba8(160, 160, 160, 255);
                        if let Some(path) = rounded_rect_path(item_x, btn_y, iw, btn_h, 2.0) {
                            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                        }
                    } else if is_hovered {
                        // Hover: light highlight with border
                        paint.set_color_rgba8(225, 230, 235, 255);
                        if let Some(path) = rounded_rect_path(item_x, btn_y, iw, btn_h, 2.0) {
                            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
                        }
                        paint.set_color_rgba8(180, 185, 190, 255);
                        if let Some(path) = rounded_rect_path(item_x, btn_y, iw, btn_h, 2.0) {
                            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                        }
                    }
                }
                ToolStripItem::Separator => {
                    // Vertical separator line
                    paint.set_color_rgba8(190, 190, 190, 255);
                    let sep_x = item_x + iw / 2.0;
                    let mut pb = PathBuilder::new();
                    pb.move_to(sep_x, btn_y + 2.0);
                    pb.line_to(sep_x, btn_y + btn_h - 2.0);
                    if let Some(path) = pb.finish() {
                        pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                    }
                }
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Hit test — returns item index of the button at position.
    pub fn hit_test(&self, mx: f32, my: f32) -> Option<usize> {
        if my < 0.0 || my > self.height { return None; }
        let positions = self.item_positions();
        for (i, (ix, iw)) in positions.iter().enumerate() {
            if matches!(self.items.get(i), Some(ToolStripItem::Button(_))) {
                if mx >= *ix && mx < ix + iw {
                    return Some(i);
                }
            }
        }
        None
    }

    /// Update hover on mouse move.
    pub fn mouse_move(&mut self, mx: f32, my: f32) {
        self.hover_index = self.hit_test(mx, my);
    }
}
