//! Button widget — standalone tiny-skia rendered button.

use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, WidgetId,
};
use super::{WidgetColors, rounded_rect_path};
use cosmic_text::Color as CosmicColor;
use tiny_skia::*;

pub struct Button {
    pub label: String,
    pub disabled: bool,
    pub pressed: bool,
    pub focused: bool,
    pub hovered: bool,
    pub colors: WidgetColors,
    pub width: f32,
    pub height: f32,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl Button {
    pub fn new(label: &str) -> Self {
        Self {
            label: label.to_string(),
            disabled: false,
            pressed: false,
            focused: false,
            hovered: false,
            colors: WidgetColors {
                background: (239, 239, 239, 255),
                border: (118, 118, 118, 255),
                ..WidgetColors::default()
            },
            width: 80.0,
            height: 28.0,
            id: WidgetId::next(),
            name: label.to_string(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Paint the button (background + border). Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background (slightly darker when pressed)
        let (r, g, b, a) = if self.pressed {
            (200, 200, 200, 255)
        } else {
            self.colors.background
        };
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 4.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Border
        let (r, g, b, a) = if self.focused {
            self.colors.focus_ring
        } else {
            self.colors.border
        };
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 4.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    pub fn click(&mut self, x: f32, y: f32) -> bool {
        if self.disabled {
            return false;
        }
        x >= 0.0 && y >= 0.0 && x <= self.width && y <= self.height
    }
}

// ── PanelWidget impl ───────────────────────────────────────────────────

impl PanelWidget for Button {
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

        // Render background + border via existing paint method
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);

        // Render label text centered
        let (fr, fg, fb, _) = self.colors.foreground;
        let font_size = 13.0;
        let text_w =
            super::ide_text::measure_text(ctx.font_system, &self.label, font_size, ctx.scale);
        let tx = r.x + (r.w - text_w) / 2.0;
        let ty = r.y + (r.h - font_size) / 2.0 - 1.0;
        super::ide_text::draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            &self.label,
            tx,
            ty,
            font_size,
            CosmicColor::rgba(fr, fg, fb, 255),
            ctx.scale,
        );
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            self.pressed = false;
            return false;
        }
        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                self.pressed = true;
                true
            }
            MouseEventKind::Release(LayoutMouseButton::Left) => {
                if self.pressed && !self.disabled {
                    self.pressed = false;
                    self.pending_events
                        .push(WidgetEvent::ButtonClicked(self.name.clone()));
                }
                true
            }
            _ => false,
        }
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused {
            return false;
        }
        use winit::event::ElementState;
        use winit::keyboard::Key;
        if event.state == ElementState::Pressed {
            if let Key::Named(winit::keyboard::NamedKey::Enter | winit::keyboard::NamedKey::Space) =
                &event.key_without_modifiers
            {
                if !self.disabled {
                    self.pending_events
                        .push(WidgetEvent::ButtonClicked(self.name.clone()));
                }
                return true;
            }
        }
        false
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetText(t) => {
                self.label = t.clone();
                CommandValue::None
            }
            WidgetCommand::GetText => CommandValue::Text(self.label.clone()),
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

    fn focusable(&self) -> bool {
        !self.disabled
    }
}
