//! MaskedTextBox widget — text field with input mask.

use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, WidgetId,
};
use super::{WidgetColors, rounded_rect_path};
use tiny_skia::*;

pub struct MaskedTextBox {
    pub mask: String,
    pub value: String,
    pub cursor: usize,
    pub focused: bool,
    pub hovered: bool,
    pub disabled: bool,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl MaskedTextBox {
    pub fn new() -> Self {
        Self {
            mask: String::new(),
            value: String::new(),
            cursor: 0,
            focused: false,
            hovered: false,
            disabled: false,
            width: 140.0,
            height: 24.0,
            colors: WidgetColors::default(),
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Paint — white background with inset border. Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // White background
        let bg = if self.disabled {
            (240, 240, 240, 255)
        } else {
            (255, 255, 255, 255)
        };
        paint.set_color_rgba8(bg.0, bg.1, bg.2, bg.3);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 2.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Inset border (sunken effect)
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        // Dark top-left
        paint.set_color_rgba8(130, 135, 144, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x, y);
        pb.line_to(x + self.width, y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Light bottom-right
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
            if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 2.0) {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

impl PanelWidget for MaskedTextBox {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }
    fn set_focused(&mut self, focused: bool) {
        self.focused = focused;
    }
    fn hovered(&self) -> bool {
        self.hovered
    }
    fn set_hovered(&mut self, hovered: bool) {
        self.hovered = hovered;
    }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.width = rect.w;
        self.height = rect.h;
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
        // Draw display text (masked)
        let display: String = if self.mask.is_empty() {
            self.value.clone()
        } else {
            self.value.chars().map(|_| '*').collect()
        };
        if !display.is_empty() {
            let (fr, fg, fb, _) = self.colors.foreground;
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &display,
                r.x + 4.0,
                r.y + 4.0,
                12.0,
                cosmic_text::Color::rgba(fr, fg, fb, 255),
                ctx.scale,
            );
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            return false;
        }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            self.focused = true;
            return true;
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused || self.disabled {
            return false;
        }
        use winit::keyboard::Key;
        match &event.logical_key {
            Key::Character(ch) => {
                self.value.push_str(ch.as_str());
                self.cursor = self.value.len();
                self.pending_events.push(WidgetEvent::TextChanged(
                    self.name.clone(),
                    self.value.clone(),
                ));
                true
            }
            Key::Named(winit::keyboard::NamedKey::Backspace) => {
                self.value.pop();
                self.cursor = self.value.len();
                self.pending_events.push(WidgetEvent::TextChanged(
                    self.name.clone(),
                    self.value.clone(),
                ));
                true
            }
            _ => false,
        }
    }

    fn focusable(&self) -> bool {
        !self.disabled
    }
    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetText(t) => {
                if self.value != *t {
                    self.value = t.clone();
                    self.pending_events.push(WidgetEvent::TextChanged(
                        self.name.clone(),
                        self.value.clone(),
                    ));
                }
                CommandValue::None
            }
            WidgetCommand::GetText => CommandValue::Text(self.value.clone()),
            WidgetCommand::GetValue => CommandValue::Text(self.value.clone()),
            WidgetCommand::SetEnabled(e) => {
                self.disabled = !e;
                CommandValue::None
            }
            _ => CommandValue::None,
        }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }
}
