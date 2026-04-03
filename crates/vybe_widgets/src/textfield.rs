//! Text input widget — standalone tiny-skia rendered text field.

use super::WidgetColors;
use cosmic_text::Color as CosmicColor;
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent};

pub struct TextInput {
    pub value: String,
    pub placeholder: String,
    pub cursor: usize,
    pub password: bool,
    pub disabled: bool,
    pub focused: bool,
    pub colors: WidgetColors,
    pub font_size: f32,
    pub width: f32,
    pub height: f32,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl TextInput {
    pub fn new() -> Self {
        Self {
            value: String::new(),
            placeholder: String::new(),
            cursor: 0,
            password: false,
            disabled: false,
            focused: false,
            colors: WidgetColors::default(),
            font_size: 14.0,
            width: 200.0,
            height: 24.0,
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_placeholder(mut self, p: &str) -> Self { self.placeholder = p.to_string(); self }
    pub fn with_password(mut self) -> Self { self.password = true; self }
    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Paint renders the border only. Text rendering requires a font system
    /// and is handled by the caller (browser engine uses cosmic_text,
    /// standalone apps can use any text renderer).
    pub fn paint_border(&self, pixmap: &mut tiny_skia::Pixmap, x: f32, y: f32, scale: f32) {
        let ts = tiny_skia::Transform::from_scale(scale, scale);
        let mut paint = tiny_skia::Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = super::rounded_rect_path(x, y, self.width, self.height, 3.0) {
            pixmap.fill_path(&path, &paint, tiny_skia::FillRule::Winding, ts, None);
        }

        // Border
        let (r, g, b, a) = if self.focused { self.colors.focus_ring } else { self.colors.border };
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = tiny_skia::Stroke::default();
        stroke.width = if self.focused { 2.0 } else { 1.0 };
        if let Some(path) = super::rounded_rect_path(x, y, self.width, self.height, 3.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    /// Get the display text (masked for password fields).
    pub fn display_text(&self) -> String {
        if self.password {
            "\u{2022}".repeat(self.value.len())
        } else if self.value.is_empty() {
            self.placeholder.clone()
        } else {
            self.value.clone()
        }
    }

    /// Is displaying placeholder text?
    pub fn is_placeholder(&self) -> bool {
        self.value.is_empty() && !self.placeholder.is_empty()
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Insert text at cursor position.
    pub fn insert(&mut self, text: &str) {
        if self.disabled { return; }
        let byte_pos = self.cursor.min(self.value.len());
        self.value.insert_str(byte_pos, text);
        self.cursor = byte_pos + text.len();
    }

    /// Delete character before cursor (backspace).
    pub fn backspace(&mut self) {
        if self.disabled || self.cursor == 0 { return; }
        let mut idx = self.cursor - 1;
        while idx > 0 && !self.value.is_char_boundary(idx) { idx -= 1; }
        self.value.drain(idx..self.cursor);
        self.cursor = idx;
    }

    /// Delete character after cursor.
    pub fn delete(&mut self) {
        if self.disabled || self.cursor >= self.value.len() { return; }
        let end = self.cursor + 1;
        let end = if end <= self.value.len() { end } else { self.value.len() };
        self.value.drain(self.cursor..end);
    }

    /// Move cursor left.
    pub fn move_left(&mut self) {
        if self.cursor > 0 {
            self.cursor -= 1;
            while self.cursor > 0 && !self.value.is_char_boundary(self.cursor) {
                self.cursor -= 1;
            }
        }
    }

    /// Move cursor right.
    pub fn move_right(&mut self) {
        if self.cursor < self.value.len() {
            self.cursor += 1;
            while self.cursor < self.value.len() && !self.value.is_char_boundary(self.cursor) {
                self.cursor += 1;
            }
        }
    }
}

// ── PanelWidget impl ───────────────────────────────────────────────────

impl PanelWidget for TextInput {
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.width = rect.w;
        self.height = rect.h;
    }

    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }

        // Border + background via existing method
        self.paint_border(ctx.pixmap, r.x, r.y, ctx.scale);

        // Text content
        let padding = 4.0;
        let display = self.display_text();
        let is_ph = self.is_placeholder();
        let (cr, cg, cb, _) = if is_ph { self.colors.placeholder } else { self.colors.foreground };
        let ty = r.y + (r.h - self.font_size) / 2.0 - 1.0;
        super::ide_text::draw_text(
            ctx.pixmap, ctx.font_system, ctx.swash_cache,
            &display, r.x + padding, ty, self.font_size,
            CosmicColor::rgba(cr, cg, cb, 255), ctx.scale,
        );

        // Cursor
        if self.focused {
            let before = if self.password {
                "\u{2022}".repeat(self.cursor.min(self.value.len()))
            } else {
                self.value[..self.cursor.min(self.value.len())].to_string()
            };
            let cursor_x = if before.is_empty() {
                0.0
            } else {
                super::ide_text::measure_text(ctx.font_system, &before, self.font_size, ctx.scale)
            };
            let cx = (r.x + padding + cursor_x) * ctx.scale;
            let cy_top = (r.y + 3.0) * ctx.scale;
            let cy_bot = (r.y + r.h - 3.0) * ctx.scale;
            let mut paint = tiny_skia::Paint::default();
            let (fr, fg, fb, _) = self.colors.foreground;
            paint.set_color_rgba8(fr, fg, fb, 255);
            let mut stroke = tiny_skia::Stroke::default();
            stroke.width = 1.0;
            let mut pb = tiny_skia::PathBuilder::new();
            pb.move_to(cx, cy_top);
            pb.line_to(cx, cy_bot);
            if let Some(path) = pb.finish() {
                ctx.pixmap.stroke_path(&path, &paint, &stroke, tiny_skia::Transform::identity(), None);
            }
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            return false;
        }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            self.focused = true;
            // Place cursor at click position
            self.cursor = self.value.len();
            return true;
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused || self.disabled { return false; }
        use winit::keyboard::{Key, NamedKey};
        use winit::event::ElementState;
        if event.state != ElementState::Pressed { return false; }

        match &event.key_without_modifiers {
            Key::Named(NamedKey::Backspace) => {
                self.backspace();
                self.pending_events.push(WidgetEvent::TextChanged(self.name.clone(), self.value.clone()));
                true
            }
            Key::Named(NamedKey::Delete) => {
                self.delete();
                self.pending_events.push(WidgetEvent::TextChanged(self.name.clone(), self.value.clone()));
                true
            }
            Key::Named(NamedKey::ArrowLeft) => { self.move_left(); true }
            Key::Named(NamedKey::ArrowRight) => { self.move_right(); true }
            Key::Named(NamedKey::Home) => { self.cursor = 0; true }
            Key::Named(NamedKey::End) => { self.cursor = self.value.len(); true }
            _ => {
                if let Some(ref text) = event.text {
                    if !text.is_empty() && text.chars().all(|c| !c.is_control()) {
                        self.insert(text);
                        self.pending_events.push(WidgetEvent::TextChanged(self.name.clone(), self.value.clone()));
                        return true;
                    }
                }
                false
            }
        }
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        if self.rect.contains(x, y) {
            winit::window::CursorIcon::Text
        } else {
            winit::window::CursorIcon::Default
        }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }

    fn focusable(&self) -> bool { !self.disabled }
}
