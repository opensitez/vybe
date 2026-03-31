//! Tab control — simple tabs header and content placeholder.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};

pub struct Tabs {
    pub tabs: Vec<String>,
    pub selected: usize,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl Tabs {
    pub fn new(labels: &[&str]) -> Self {
        Self {
            tabs: labels.iter().map(|s| s.to_string()).collect(),
            selected: 0,
            width: 300.0,
            height: 200.0,
            colors: WidgetColors::default(),
        }
    }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        let header_h = 28.0;
        let tab_w = (self.width / self.tabs.len() as f32).max(60.0);

        // Background for header
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = rounded_rect_path(x, y, self.width, header_h, 4.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Tabs
        for (i, label) in self.tabs.iter().enumerate() {
            let tx = x + i as f32 * tab_w;
            let tw = tab_w;
            let is_sel = i == self.selected;
            // Tab background
            if is_sel {
                let (r, g, b, a) = self.colors.accent;
                paint.set_color_rgba8(r, g, b, a);
            } else {
                let (r, g, b, a) = (240, 240, 240, 255);
                paint.set_color_rgba8(r, g, b, a);
            }
            if let Some(path) = rounded_rect_path(tx + 4.0, y + 4.0, tw - 8.0, header_h - 8.0, 3.0) {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }
            // Tab border
            let (br, bg, bb, ba) = self.colors.border;
            paint.set_color_rgba8(br, bg, bb, ba);
            let mut stroke = Stroke::default();
            stroke.width = 1.0;
            if let Some(path) = rounded_rect_path(tx + 4.0, y + 4.0, tw - 8.0, header_h - 8.0, 3.0) {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
            // Label (simple mono placeholder using border color as no font engine here)
            // For demos the existing demo harness uses cosmic-text; here we leave content blank.
        }

        // Content area
        let content_y = y + header_h;
        let content_h = self.height - header_h;
        let (cr, cg, cb, ca) = self.colors.background;
        paint.set_color_rgba8(cr, cg, cb, ca);
        if let Some(path) = rounded_rect_path(x, content_y, self.width, content_h, 4.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }
        // Border
        let (br, bg, bb, ba) = self.colors.border;
        paint.set_color_rgba8(br, bg, bb, ba);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        if let Some(path) = rounded_rect_path(x, content_y, self.width, content_h, 4.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) { (self.width, self.height) }

    pub fn click(&mut self, mx: f32, my: f32, x: f32, y: f32) -> bool {
        let header_h = 28.0;
        if my < y || my > y + header_h { return false; }
        let tab_w = (self.width / self.tabs.len() as f32).max(60.0);
        let idx = ((mx - x) / tab_w).floor() as isize;
        if idx >= 0 && (idx as usize) < self.tabs.len() {
            self.selected = idx as usize;
            return true;
        }
        false
    }
}
