//! DateTimePicker widget — text area with dropdown button.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent};

pub struct DateTimePicker {
    pub value: String,
    pub focused: bool,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub name: String,
    rect: LayoutRect,
}

impl DateTimePicker {
    pub fn new() -> Self {
        Self {
            value: String::new(),
            focused: false,
            width: 140.0,
            height: 24.0,
            colors: WidgetColors::default(),
            name: String::new(),
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Width of the dropdown button area.
    fn button_width(&self) -> f32 {
        20.0
    }

    /// Paint the date time picker — white text area + dropdown button.
    /// Date text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let bw = self.button_width();

        // White text area background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 1.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Dropdown button background
        paint.set_color_rgba8(240, 240, 240, 255);
        let btn_x = x + self.width - bw;
        if let Some(rect) = Rect::from_xywh(btn_x, y + 1.0, bw - 1.0, self.height - 2.0) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Divider
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        paint.set_color_rgba8(160, 160, 160, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(btn_x, y + 1.0);
        pb.line_to(btn_x, y + self.height - 1.0);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Dropdown arrow (downward pointing triangle)
        let arrow_size = 4.0;
        let center_x = btn_x + bw / 2.0;
        let center_y = y + self.height / 2.0;
        paint.set_color_rgba8(60, 60, 60, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(center_x, center_y + arrow_size);
        pb.line_to(center_x + arrow_size, center_y - arrow_size * 0.5);
        pb.line_to(center_x - arrow_size, center_y - arrow_size * 0.5);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Outer border
        let (r, g, b, a) = if self.focused { self.colors.focus_ring } else { self.colors.border };
        paint.set_color_rgba8(r, g, b, a);
        stroke.width = 1.0;
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 1.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Returns true if click is on the dropdown button.
    pub fn click_dropdown(&self, click_x: f32, _click_y: f32) -> bool {
        click_x >= self.width - self.button_width()
    }
}

impl PanelWidget for DateTimePicker {
    fn name(&self) -> &str { &self.name }
    fn set_focused(&mut self, focused: bool) { self.focused = focused; }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; self.width = rect.w; self.height = rect.h; }
    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
        if !self.value.is_empty() {
            let (fr, fg, fb, _) = self.colors.foreground;
            super::ide_text::draw_text(ctx.pixmap, ctx.font_system, ctx.swash_cache, &self.value, r.x + 4.0, r.y + 4.0, 12.0, cosmic_text::Color::rgba(fr, fg, fb, 255), ctx.scale);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) { return false; }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            self.focused = true;
            return true;
        }
        false
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
    fn focusable(&self) -> bool { true }
}
