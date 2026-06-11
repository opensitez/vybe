//! Range slider widget — standalone tiny-skia rendered slider.

use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, WidgetId,
};
use super::{WidgetColors, circle_path, rounded_rect_path};
use tiny_skia::*;

pub struct Slider {
    pub value: f32, // 0.0..1.0
    pub min: f32,
    pub max: f32,
    pub disabled: bool,
    pub focused: bool,
    pub hovered: bool,
    pub dragging: bool,
    pub colors: WidgetColors,
    pub width: f32,
    pub height: f32,
    pub track_height: f32,
    pub thumb_radius: f32,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl Slider {
    pub fn new(min: f32, max: f32, value: f32) -> Self {
        let pct = if max > min {
            (value - min) / (max - min)
        } else {
            0.0
        };
        Self {
            value: pct.clamp(0.0, 1.0),
            min,
            max,
            disabled: false,
            focused: false,
            hovered: false,
            dragging: false,
            colors: WidgetColors::default(),
            width: 200.0,
            height: 20.0,
            track_height: 4.0,
            thumb_radius: 8.0,
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

    /// Get the actual value (mapped from 0..1 to min..max).
    pub fn actual_value(&self) -> f32 {
        self.min + self.value * (self.max - self.min)
    }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        let track_y = y + (self.height - self.track_height) / 2.0;
        let thumb_x = x + self.thumb_radius + self.value * (self.width - self.thumb_radius * 2.0);
        let thumb_y = y + self.height / 2.0;

        // Track background
        paint.set_color_rgba8(200, 200, 200, 255);
        if let Some(path) = rounded_rect_path(x, track_y, self.width, self.track_height, 2.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Filled portion
        let (r, g, b, a) = self.colors.accent;
        paint.set_color_rgba8(r, g, b, a);
        let fill_w = thumb_x - x;
        if fill_w > 0.0 {
            if let Some(path) = rounded_rect_path(x, track_y, fill_w, self.track_height, 2.0) {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }
        }

        // Thumb
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = circle_path(thumb_x, thumb_y, self.thumb_radius) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }
        let (r, g, b, a) = if self.focused {
            self.colors.focus_ring
        } else {
            self.colors.border
        };
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        if let Some(path) = circle_path(thumb_x, thumb_y, self.thumb_radius) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Handle mouse down — start dragging if on thumb.
    pub fn mouse_down(&mut self, x: f32, _y: f32) -> bool {
        if self.disabled {
            return false;
        }
        self.dragging = true;
        self.set_from_x(x);
        true
    }

    /// Handle mouse move during drag.
    pub fn mouse_move(&mut self, x: f32) {
        if self.dragging {
            self.set_from_x(x);
        }
    }

    /// Handle mouse up — stop dragging.
    pub fn mouse_up(&mut self) {
        self.dragging = false;
    }

    fn set_from_x(&mut self, x: f32) {
        let usable = self.width - self.thumb_radius * 2.0;
        let pct = (x - self.thumb_radius) / usable;
        self.value = pct.clamp(0.0, 1.0);
    }
}

impl PanelWidget for Slider {
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
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        let lx = event.x - r.x;
        let ly = event.y - r.y;
        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                if r.contains(event.x, event.y) {
                    self.mouse_down(lx, ly);
                    self.pending_events.push(WidgetEvent::SliderChanged(
                        self.name.clone(),
                        self.actual_value(),
                    ));
                    return true;
                }
            }
            MouseEventKind::Move => {
                if self.dragging {
                    self.mouse_move(lx);
                    self.pending_events.push(WidgetEvent::SliderChanged(
                        self.name.clone(),
                        self.actual_value(),
                    ));
                    return true;
                }
            }
            MouseEventKind::Release(LayoutMouseButton::Left) => {
                if self.dragging {
                    self.mouse_up();
                    return true;
                }
            }
            _ => {}
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused {
            return false;
        }
        use winit::keyboard::{Key, NamedKey};
        match &event.logical_key {
            Key::Named(NamedKey::ArrowRight) | Key::Named(NamedKey::ArrowUp) => {
                self.value = (self.value + 0.05).clamp(0.0, 1.0);
                self.pending_events.push(WidgetEvent::SliderChanged(
                    self.name.clone(),
                    self.actual_value(),
                ));
                true
            }
            Key::Named(NamedKey::ArrowLeft) | Key::Named(NamedKey::ArrowDown) => {
                self.value = (self.value - 0.05).clamp(0.0, 1.0);
                self.pending_events.push(WidgetEvent::SliderChanged(
                    self.name.clone(),
                    self.actual_value(),
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
            WidgetCommand::SetValue(v) => {
                let new_v = *v as f32;
                if (self.value - new_v).abs() > f32::EPSILON {
                    self.value = new_v;
                    self.pending_events.push(WidgetEvent::SliderChanged(
                        self.name.clone(),
                        self.actual_value(),
                    ));
                }
                CommandValue::None
            }
            WidgetCommand::GetValue => CommandValue::Number(self.value as f64),
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

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        if self.rect.contains(x, y) {
            winit::window::CursorIcon::Pointer
        } else {
            winit::window::CursorIcon::Default
        }
    }
}
