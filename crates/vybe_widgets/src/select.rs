//! Select/dropdown widget — standalone tiny-skia rendered select box.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, WidgetId,
};
use tiny_skia::*;

pub struct Select {
    pub options: Vec<String>,
    pub selected_index: usize,
    pub open: bool,
    pub disabled: bool,
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

impl Select {
    pub fn new(options: Vec<String>) -> Self {
        Self {
            options,
            selected_index: 0,
            open: false,
            disabled: false,
            focused: false,
            hovered: false,
            colors: WidgetColors::default(),
            width: 200.0,
            height: 24.0,
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

    pub fn selected_text(&self) -> &str {
        self.options
            .get(self.selected_index)
            .map(|s| s.as_str())
            .unwrap_or("")
    }

    /// Paint the select box (closed state). Draws border + dropdown arrow.
    /// Text rendering is handled by the caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = super::rounded_rect_path(x, y, self.width, self.height, 3.0) {
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
        stroke.width = if self.focused { 2.0 } else { 1.0 };
        if let Some(path) = super::rounded_rect_path(x, y, self.width, self.height, 3.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Dropdown arrow
        let arrow_x = x + self.width - 16.0;
        let arrow_y = y + self.height / 2.0;
        let (r, g, b, a) = self.colors.foreground;
        paint.set_color_rgba8(r, g, b, a);
        stroke.width = 1.5;
        let mut pb = PathBuilder::new();
        pb.move_to(arrow_x - 4.0, arrow_y - 2.0);
        pb.line_to(arrow_x, arrow_y + 2.0);
        pb.line_to(arrow_x + 4.0, arrow_y - 2.0);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    pub fn click(&mut self, _x: f32, _y: f32) -> bool {
        if self.disabled {
            return false;
        }
        self.open = !self.open;
        true
    }

    pub fn select_index(&mut self, idx: usize) {
        if idx < self.options.len() {
            self.selected_index = idx;
            self.open = false;
        }
    }
}

impl PanelWidget for Select {
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
        // Draw selected text
        let txt = self.selected_text();
        if !txt.is_empty() {
            let (fr, fg, fb, _) = self.colors.foreground;
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                txt,
                r.x + 6.0,
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
            self.click(event.x - self.rect.x, event.y - self.rect.y);
            self.pending_events.push(WidgetEvent::SelectChanged(
                self.name.clone(),
                self.selected_index,
            ));
            return true;
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused {
            return false;
        }
        use winit::keyboard::{Key, NamedKey};
        match &event.logical_key {
            Key::Named(NamedKey::ArrowDown) => {
                if self.selected_index + 1 < self.options.len() {
                    self.selected_index += 1;
                    self.pending_events.push(WidgetEvent::SelectChanged(
                        self.name.clone(),
                        self.selected_index,
                    ));
                }
                true
            }
            Key::Named(NamedKey::ArrowUp) => {
                if self.selected_index > 0 {
                    self.selected_index -= 1;
                    self.pending_events.push(WidgetEvent::SelectChanged(
                        self.name.clone(),
                        self.selected_index,
                    ));
                }
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
            WidgetCommand::SetSelectedIndex(i) => {
                if *i < self.options.len() && self.selected_index != *i {
                    self.select_index(*i);
                    self.pending_events.push(WidgetEvent::SelectChanged(
                        self.name.clone(),
                        self.selected_index,
                    ));
                }
                CommandValue::None
            }
            WidgetCommand::GetValue => CommandValue::Index(self.selected_index),
            WidgetCommand::AddItem(s) => {
                self.options.push(s.clone());
                CommandValue::None
            }
            WidgetCommand::GetItem(i) => match self.options.get(*i) {
                Some(text) => CommandValue::Text(text.clone()),
                None => CommandValue::None,
            },
            WidgetCommand::SetItem(i, text) => {
                if let Some(slot) = self.options.get_mut(*i) {
                    *slot = text.clone();
                }
                CommandValue::None
            }
            WidgetCommand::RemoveItem(i) => {
                if *i < self.options.len() {
                    self.options.remove(*i);
                }
                CommandValue::None
            }
            WidgetCommand::ClearItems => {
                self.options.clear();
                self.selected_index = 0;
                CommandValue::None
            }
            WidgetCommand::SetEnabled(e) => {
                self.disabled = !e;
                CommandValue::None
            }
            WidgetCommand::GetText => {
                let t = self
                    .options
                    .get(self.selected_index)
                    .cloned()
                    .unwrap_or_default();
                CommandValue::Text(t)
            }
            _ => CommandValue::None,
        }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }
}
