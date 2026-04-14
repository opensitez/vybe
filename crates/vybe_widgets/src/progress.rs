//! Progress bar widget

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};
use super::layout::{LayoutRect, MouseEvent, KeyEvent, RenderContext, PanelWidget, WidgetId, WidgetCommand, CommandValue};

pub struct ProgressBar {
    pub value: f32, // 0.0..1.0
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
}

impl ProgressBar {
    pub fn new() -> Self {
        Self { value: 0.0, width: 200.0, height: 16.0, colors: WidgetColors::default(), id: WidgetId::next(), name: String::new(), rect: LayoutRect::zero() }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = (240, 240, 240, 255);
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 4.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Filled portion
        let fill_w = (self.value.clamp(0.0, 1.0)) * self.width;
        let (fr, fg, fb, fa) = self.colors.accent;
        paint.set_color_rgba8(fr, fg, fb, fa);
        if fill_w > 0.0 {
            if let Some(path) = rounded_rect_path(x, y, fill_w, self.height, 4.0) {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }
        }

        // Border
        let (br, bg, bb, ba) = self.colors.border;
        paint.set_color_rgba8(br, bg, bb, ba);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 4.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) { (self.width, self.height) }
}

impl PanelWidget for ProgressBar {
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

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetValue(v) => { self.value = *v as f32; CommandValue::None }
            WidgetCommand::GetValue => CommandValue::Number(self.value as f64),
            _ => CommandValue::None,
        }
    }
}
