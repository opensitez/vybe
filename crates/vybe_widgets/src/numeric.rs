//! NumericUpDown widget — text field with spinner buttons.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent, WidgetId};

pub struct NumericUpDown {
    pub value: f64,
    pub min: f64,
    pub max: f64,
    pub increment: f64,
    pub focused: bool,
    pub hovered: bool,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl NumericUpDown {
    pub fn new() -> Self {
        Self {
            value: 0.0,
            min: 0.0,
            max: 100.0,
            increment: 1.0,
            focused: false,
            hovered: false,
            width: 80.0,
            height: 24.0,
            colors: WidgetColors::default(),
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Width of the spinner button area.
    fn button_width(&self) -> f32 {
        17.0
    }

    /// Paint the numeric up-down — white text area + up/down spinner buttons.
    /// Value text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let bw = self.button_width();

        // White text field background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 1.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Spinner button background
        paint.set_color_rgba8(240, 240, 240, 255);
        let btn_x = x + self.width - bw;
        if let Some(rect) = Rect::from_xywh(btn_x, y + 1.0, bw - 1.0, self.height - 2.0) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Divider between text and buttons
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        paint.set_color_rgba8(160, 160, 160, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(btn_x, y + 1.0);
        pb.line_to(btn_x, y + self.height - 1.0);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Horizontal divider between up and down buttons
        let mid_y = y + self.height / 2.0;
        let mut pb = PathBuilder::new();
        pb.move_to(btn_x, mid_y);
        pb.line_to(x + self.width - 1.0, mid_y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Up triangle
        let arrow_size = 3.5;
        let center_x = btn_x + bw / 2.0;
        let up_center_y = y + self.height * 0.25;
        paint.set_color_rgba8(60, 60, 60, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(center_x, up_center_y - arrow_size);
        pb.line_to(center_x + arrow_size, up_center_y + arrow_size * 0.6);
        pb.line_to(center_x - arrow_size, up_center_y + arrow_size * 0.6);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Down triangle
        let down_center_y = y + self.height * 0.75;
        let mut pb = PathBuilder::new();
        pb.move_to(center_x, down_center_y + arrow_size);
        pb.line_to(center_x + arrow_size, down_center_y - arrow_size * 0.6);
        pb.line_to(center_x - arrow_size, down_center_y - arrow_size * 0.6);
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

    /// Increment the value (clamped to max).
    pub fn increment(&mut self) {
        self.value = (self.value + self.increment).min(self.max);
    }

    /// Decrement the value (clamped to min).
    pub fn decrement(&mut self) {
        self.value = (self.value - self.increment).max(self.min);
    }

    /// Handle click — returns true if value changed. Up button is top half of
    /// spinner area, down button is bottom half.
    pub fn click(&mut self, click_x: f32, click_y: f32) -> bool {
        let bw = self.button_width();
        let btn_x = self.width - bw;
        if click_x >= btn_x && click_x <= self.width && click_y >= 0.0 && click_y <= self.height {
            if click_y < self.height / 2.0 {
                self.increment();
            } else {
                self.decrement();
            }
            return true;
        }
        false
    }

    /// Display text for the value.
    pub fn display_text(&self) -> String {
        if self.value == self.value.floor() {
            format!("{}", self.value as i64)
        } else {
            format!("{}", self.value)
        }
    }
}

impl PanelWidget for NumericUpDown {
    fn name(&self) -> &str { &self.name }
    fn widget_id(&self) -> WidgetId { self.id }
    fn set_focused(&mut self, focused: bool) { self.focused = focused; }
    fn hovered(&self) -> bool { self.hovered }
    fn set_hovered(&mut self, hovered: bool) { self.hovered = hovered; }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; self.width = rect.w; self.height = rect.h; }
    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
        let txt = self.display_text();
        let (fr, fg, fb, _) = self.colors.foreground;
        super::ide_text::draw_text(ctx.pixmap, ctx.font_system, ctx.swash_cache, &txt, r.x + 4.0, r.y + 4.0, 12.0, cosmic_text::Color::rgba(fr, fg, fb, 255), ctx.scale);
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        if !r.contains(event.x, event.y) { return false; }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            if self.click(event.x - r.x, event.y - r.y) {
                self.pending_events.push(WidgetEvent::NumericChanged(self.name.clone(), self.value));
                return true;
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused { return false; }
        use winit::keyboard::{Key, NamedKey};
        match &event.logical_key {
            Key::Named(NamedKey::ArrowUp) => { self.increment(); self.pending_events.push(WidgetEvent::NumericChanged(self.name.clone(), self.value)); true }
            Key::Named(NamedKey::ArrowDown) => { self.decrement(); self.pending_events.push(WidgetEvent::NumericChanged(self.name.clone(), self.value)); true }
            _ => false,
        }
    }

    fn focusable(&self) -> bool { true }
    fn drain_events(&mut self) -> Vec<WidgetEvent> { std::mem::take(&mut self.pending_events) }
}
