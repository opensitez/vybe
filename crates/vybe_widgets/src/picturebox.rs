//! PictureBox widget — image placeholder with X cross and icon.

use tiny_skia::*;
use super::WidgetColors;
use super::layout::{LayoutRect, MouseEvent, KeyEvent, RenderContext, PanelWidget, WidgetEvent, WidgetId, WidgetCommand, CommandValue};

pub struct PictureBox {
    pub has_image: bool,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
}

impl PictureBox {
    pub fn new() -> Self {
        Self {
            has_image: false,
            width: 160.0,
            height: 120.0,
            colors: WidgetColors::default(),
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Paint — gray background with X cross lines as placeholder and a small image icon.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Gray background
        paint.set_color_rgba8(240, 240, 240, 255);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Border
        paint.set_color_rgba8(200, 200, 200, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        if !self.has_image {
            // Placeholder X cross lines
            paint.set_color_rgba8(210, 210, 210, 255);
            stroke.width = 1.0;

            // Top-left to bottom-right diagonal
            let mut pb = PathBuilder::new();
            pb.move_to(x + 2.0, y + 2.0);
            pb.line_to(x + self.width - 2.0, y + self.height - 2.0);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Top-right to bottom-left diagonal
            let mut pb = PathBuilder::new();
            pb.move_to(x + self.width - 2.0, y + 2.0);
            pb.line_to(x + 2.0, y + self.height - 2.0);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Small image icon in center (mountain/landscape symbol)
            let cx = x + self.width / 2.0;
            let cy = y + self.height / 2.0;
            let icon_s = 12.0;
            paint.set_color_rgba8(180, 180, 180, 255);
            stroke.width = 1.5;

            // Image frame
            let mut pb = PathBuilder::new();
            pb.move_to(cx - icon_s, cy - icon_s * 0.7);
            pb.line_to(cx + icon_s, cy - icon_s * 0.7);
            pb.line_to(cx + icon_s, cy + icon_s * 0.7);
            pb.line_to(cx - icon_s, cy + icon_s * 0.7);
            pb.close();
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Mountain shape inside frame
            paint.set_color_rgba8(160, 160, 160, 255);
            stroke.width = 1.0;
            let mut pb = PathBuilder::new();
            pb.move_to(cx - icon_s + 2.0, cy + icon_s * 0.5);
            pb.line_to(cx - icon_s * 0.3, cy - icon_s * 0.2);
            pb.line_to(cx, cy + icon_s * 0.1);
            pb.line_to(cx + icon_s * 0.3, cy - icon_s * 0.4);
            pb.line_to(cx + icon_s - 2.0, cy + icon_s * 0.5);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Sun circle
            if let Some(path) = super::circle_path(cx - icon_s * 0.5, cy - icon_s * 0.3, 2.5) {
                paint.set_color_rgba8(180, 180, 180, 255);
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

impl PanelWidget for PictureBox {
    fn name(&self) -> &str { &self.name }
    fn widget_id(&self) -> WidgetId { self.id }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; self.width = rect.w; self.height = rect.h; }
    fn rect(&self) -> LayoutRect { self.rect }
    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
    }
    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool { false }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
}
