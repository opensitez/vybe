//! Text input widget — standalone tiny-skia rendered text field.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, WidgetCommand, WidgetEvent, WidgetId,
};
use cosmic_text::Color as CosmicColor;

pub struct TextInput {
    pub value: String,
    pub placeholder: String,
    pub cursor: usize,
    /// Selection anchor (the other end of the selection range). `None` = no selection.
    pub selection_anchor: Option<usize>,
    pub password: bool,
    pub disabled: bool,
    pub read_only: bool,
    pub max_length: Option<usize>,
    pub focused: bool,
    pub hovered: bool,
    pub colors: WidgetColors,
    pub font_size: f32,
    pub width: f32,
    pub height: f32,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
    /// Whether a mouse drag selection is in progress.
    dragging: bool,
}

impl TextInput {
    pub fn new() -> Self {
        Self {
            value: String::new(),
            placeholder: String::new(),
            cursor: 0,
            selection_anchor: None,
            password: false,
            disabled: false,
            read_only: false,
            max_length: None,
            focused: false,
            hovered: false,
            colors: WidgetColors::default(),
            font_size: 14.0,
            width: 200.0,
            height: 24.0,
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
            dragging: false,
        }
    }

    pub fn with_placeholder(mut self, p: &str) -> Self {
        self.placeholder = p.to_string();
        self
    }
    pub fn with_password(mut self) -> Self {
        self.password = true;
        self
    }
    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }
    pub fn with_read_only(mut self) -> Self {
        self.read_only = true;
        self
    }
    pub fn with_max_length(mut self, max: usize) -> Self {
        self.max_length = Some(max);
        self
    }

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
        let (r, g, b, a) = if self.focused {
            self.colors.focus_ring
        } else {
            self.colors.border
        };
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

    /// Insert text at cursor position (replaces selection if any).
    pub fn insert(&mut self, text: &str) {
        if self.disabled || self.read_only {
            return;
        }
        self.delete_selection();
        let insert_text = if let Some(max) = self.max_length {
            let remaining = max.saturating_sub(self.value.chars().count());
            let t: String = text.chars().take(remaining).collect();
            t
        } else {
            text.to_string()
        };
        let byte_pos = self.cursor.min(self.value.len());
        self.value.insert_str(byte_pos, &insert_text);
        self.cursor = byte_pos + insert_text.len();
    }

    /// Delete character before cursor (backspace). If selection active, delete selection.
    pub fn backspace(&mut self) {
        if self.disabled || self.read_only {
            return;
        }
        if self.has_selection() {
            self.delete_selection();
            return;
        }
        if self.cursor == 0 {
            return;
        }
        let mut idx = self.cursor - 1;
        while idx > 0 && !self.value.is_char_boundary(idx) {
            idx -= 1;
        }
        self.value.drain(idx..self.cursor);
        self.cursor = idx;
    }

    /// Delete character after cursor. If selection active, delete selection.
    pub fn delete(&mut self) {
        if self.disabled || self.read_only {
            return;
        }
        if self.has_selection() {
            self.delete_selection();
            return;
        }
        if self.cursor >= self.value.len() {
            return;
        }
        let end = self.cursor + 1;
        let end = if end <= self.value.len() {
            end
        } else {
            self.value.len()
        };
        self.value.drain(self.cursor..end);
    }

    /// Returns the selection range (start_byte, end_byte) sorted, or None.
    pub fn selection_range(&self) -> Option<(usize, usize)> {
        self.selection_anchor.map(|anchor| {
            let a = anchor.min(self.value.len());
            let b = self.cursor.min(self.value.len());
            if a <= b { (a, b) } else { (b, a) }
        })
    }

    /// Whether there is a non-empty selection.
    pub fn has_selection(&self) -> bool {
        self.selection_range().map_or(false, |(a, b)| a != b)
    }

    /// Get the selected text.
    pub fn selected_text(&self) -> String {
        if let Some((start, end)) = self.selection_range() {
            self.value[start..end].to_string()
        } else {
            String::new()
        }
    }

    /// Select all text.
    pub fn select_all(&mut self) {
        if self.value.is_empty() {
            return;
        }
        self.selection_anchor = Some(0);
        self.cursor = self.value.len();
    }

    /// Clear the selection without deleting.
    pub fn clear_selection(&mut self) {
        self.selection_anchor = None;
    }

    /// Delete the selected text and place cursor at start of selection.
    pub fn delete_selection(&mut self) {
        if let Some((start, end)) = self.selection_range() {
            if start != end {
                self.value.drain(start..end);
                self.cursor = start;
            }
        }
        self.selection_anchor = None;
    }

    /// Hit-test: given an x position (in widget-local coords), return byte offset in value.
    #[allow(dead_code)]
    fn hit_test_cursor(
        &self,
        font_system: &mut cosmic_text::FontSystem,
        local_x: f32,
        scale: f32,
    ) -> usize {
        let text = if self.password {
            "\u{2022}".repeat(self.value.len())
        } else {
            self.value.clone()
        };
        if text.is_empty() {
            return 0;
        }
        // Find the byte position whose rendered width is closest to local_x
        let mut best_pos = 0;
        let mut best_dist = local_x.abs();
        let mut pos = 0;
        for ch in text.chars() {
            pos += ch.len_utf8();
            let w = super::ide_text::measure_text(font_system, &text[..pos], self.font_size, scale);
            let dist = (w - local_x).abs();
            if dist < best_dist {
                best_dist = dist;
                best_pos = pos;
            }
        }
        // For password mode, map back to value byte offset (same since bullet is multi-byte but
        // we used value.len() bullets — the char count matches)
        if self.password {
            // Each bullet is 3 bytes in UTF-8; map back to character index
            let char_idx = best_pos / "\u{2022}".len();
            self.value
                .char_indices()
                .nth(char_idx)
                .map_or(self.value.len(), |(i, _)| i)
        } else {
            best_pos.min(self.value.len())
        }
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

        // Border + background via existing method
        self.paint_border(ctx.pixmap, r.x, r.y, ctx.scale);

        // Text content
        let padding = 4.0;
        let display = self.display_text();
        let is_ph = self.is_placeholder();
        let (cr, cg, cb, _) = if is_ph {
            self.colors.placeholder
        } else {
            self.colors.foreground
        };
        let ty = r.y + (r.h - self.font_size) / 2.0 - 1.0;

        // Selection highlight
        if self.focused {
            if let Some((sel_start, sel_end)) = self.selection_range() {
                if sel_start != sel_end {
                    let sel_display_start = if self.password {
                        "\u{2022}".repeat(self.value[..sel_start].chars().count())
                    } else {
                        self.value[..sel_start].to_string()
                    };
                    let sel_display_end = if self.password {
                        "\u{2022}".repeat(self.value[..sel_end].chars().count())
                    } else {
                        self.value[..sel_end].to_string()
                    };
                    let x_start = if sel_display_start.is_empty() {
                        0.0
                    } else {
                        super::ide_text::measure_text(
                            ctx.font_system,
                            &sel_display_start,
                            self.font_size,
                            ctx.scale,
                        )
                    };
                    let x_end = if sel_display_end.is_empty() {
                        0.0
                    } else {
                        super::ide_text::measure_text(
                            ctx.font_system,
                            &sel_display_end,
                            self.font_size,
                            ctx.scale,
                        )
                    };
                    // Draw selection rectangle
                    let sx = (r.x + padding + x_start) * ctx.scale;
                    let sw = (x_end - x_start) * ctx.scale;
                    let sy = (r.y + 2.0) * ctx.scale;
                    let sh = (r.h - 4.0) * ctx.scale;
                    if let Some(rect) = tiny_skia::Rect::from_xywh(sx, sy, sw.max(1.0), sh) {
                        let mut paint = tiny_skia::Paint::default();
                        paint.set_color_rgba8(51, 153, 255, 100); // blue selection highlight
                        ctx.pixmap
                            .fill_rect(rect, &paint, tiny_skia::Transform::identity(), None);
                    }
                }
            }
        }

        super::ide_text::draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            &display,
            r.x + padding,
            ty,
            self.font_size,
            CosmicColor::rgba(cr, cg, cb, 255),
            ctx.scale,
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
                ctx.pixmap.stroke_path(
                    &path,
                    &paint,
                    &stroke,
                    tiny_skia::Transform::identity(),
                    None,
                );
            }
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            self.dragging = false;
            return false;
        }
        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                self.focused = true;
                // Need font system for hit-test; approximate with simple char-width method
                // A proper hit-test is done in render with measure_text; here we store basic position.
                // We'll refine in render, but for click we set the flag:
                let padding = 4.0;
                let local_x = event.x - self.rect.x - padding;
                // Approximate: set cursor to end; proper hit-test happens via `hit_test_cursor` in handle_mouse_with_font.
                // For now, use a simple proportional approximation.
                if self.value.is_empty() {
                    self.cursor = 0;
                } else {
                    let char_count = self.value.chars().count();
                    let avg_char_width = self.width / char_count.max(1) as f32;
                    let char_idx = ((local_x / avg_char_width).round() as usize).min(char_count);
                    // Convert char index to byte position
                    self.cursor = self
                        .value
                        .char_indices()
                        .nth(char_idx)
                        .map_or(self.value.len(), |(i, _)| i);
                }
                if event.shift {
                    // Shift+click: extend selection
                    if self.selection_anchor.is_none() {
                        self.selection_anchor = Some(self.cursor);
                    }
                } else {
                    self.selection_anchor = None;
                }
                self.dragging = true;
                return true;
            }
            MouseEventKind::Move => {
                if self.dragging {
                    let padding = 4.0;
                    let local_x = event.x - self.rect.x - padding;
                    if self.selection_anchor.is_none() {
                        self.selection_anchor = Some(self.cursor);
                    }
                    if self.value.is_empty() {
                        self.cursor = 0;
                    } else {
                        let char_count = self.value.chars().count();
                        let avg_char_width = self.width / char_count.max(1) as f32;
                        let char_idx =
                            ((local_x / avg_char_width).round() as usize).min(char_count);
                        self.cursor = self
                            .value
                            .char_indices()
                            .nth(char_idx)
                            .map_or(self.value.len(), |(i, _)| i);
                    }
                    return true;
                }
            }
            MouseEventKind::Release(LayoutMouseButton::Left) => {
                self.dragging = false;
            }
            _ => {}
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused || self.disabled {
            return false;
        }
        use winit::event::ElementState;
        use winit::keyboard::{Key, NamedKey};
        if event.state != ElementState::Pressed {
            return false;
        }

        let is_shift = event.shift;
        let is_cmd = event.cmd; // Cmd on macOS, Ctrl on others

        // --- Clipboard & select-all shortcuts ---
        if is_cmd {
            match &event.key_without_modifiers {
                Key::Character(c) if c.as_str() == "a" => {
                    self.select_all();
                    return true;
                }
                Key::Character(c) if c.as_str() == "c" => {
                    if self.has_selection() {
                        if let Ok(mut cb) = arboard::Clipboard::new() {
                            let _ = cb.set_text(self.selected_text());
                        }
                    }
                    return true;
                }
                Key::Character(c) if c.as_str() == "v" => {
                    if !self.read_only {
                        if let Ok(mut cb) = arboard::Clipboard::new() {
                            if let Ok(text) = cb.get_text() {
                                self.insert(&text);
                                self.pending_events.push(WidgetEvent::TextChanged(
                                    self.name.clone(),
                                    self.value.clone(),
                                ));
                            }
                        }
                    }
                    return true;
                }
                Key::Character(c) if c.as_str() == "x" => {
                    if self.has_selection() {
                        if let Ok(mut cb) = arboard::Clipboard::new() {
                            let _ = cb.set_text(self.selected_text());
                        }
                        if !self.read_only {
                            self.delete_selection();
                            self.pending_events.push(WidgetEvent::TextChanged(
                                self.name.clone(),
                                self.value.clone(),
                            ));
                        }
                    }
                    return true;
                }
                _ => {}
            }
        }

        match &event.key_without_modifiers {
            Key::Named(NamedKey::Backspace) => {
                self.backspace();
                self.pending_events.push(WidgetEvent::TextChanged(
                    self.name.clone(),
                    self.value.clone(),
                ));
                true
            }
            Key::Named(NamedKey::Delete) => {
                self.delete();
                self.pending_events.push(WidgetEvent::TextChanged(
                    self.name.clone(),
                    self.value.clone(),
                ));
                true
            }
            Key::Named(NamedKey::ArrowLeft) => {
                if is_shift {
                    if self.selection_anchor.is_none() {
                        self.selection_anchor = Some(self.cursor);
                    }
                    self.move_left();
                } else {
                    if self.has_selection() {
                        let (start, _) = self.selection_range().unwrap();
                        self.cursor = start;
                        self.clear_selection();
                    } else {
                        self.move_left();
                    }
                }
                true
            }
            Key::Named(NamedKey::ArrowRight) => {
                if is_shift {
                    if self.selection_anchor.is_none() {
                        self.selection_anchor = Some(self.cursor);
                    }
                    self.move_right();
                } else {
                    if self.has_selection() {
                        let (_, end) = self.selection_range().unwrap();
                        self.cursor = end;
                        self.clear_selection();
                    } else {
                        self.move_right();
                    }
                }
                true
            }
            Key::Named(NamedKey::Home) => {
                if is_shift {
                    if self.selection_anchor.is_none() {
                        self.selection_anchor = Some(self.cursor);
                    }
                } else {
                    self.clear_selection();
                }
                self.cursor = 0;
                true
            }
            Key::Named(NamedKey::End) => {
                if is_shift {
                    if self.selection_anchor.is_none() {
                        self.selection_anchor = Some(self.cursor);
                    }
                } else {
                    self.clear_selection();
                }
                self.cursor = self.value.len();
                true
            }
            _ => {
                if let Some(ref text) = event.text {
                    if !text.is_empty() && text.chars().all(|c| !c.is_control()) {
                        self.insert(text);
                        self.pending_events.push(WidgetEvent::TextChanged(
                            self.name.clone(),
                            self.value.clone(),
                        ));
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

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetText(t) => {
                // Programmatic write. Fire TextChanged on actual value
                // change so user-installed handlers see it the same way
                // they see typed input. Matches the path in
                // `handle_key` / `handle_mouse` which also push
                // TextChanged when the value mutates.
                if self.value != *t {
                    self.value = t.clone();
                    self.cursor = self.value.len();
                    self.selection_anchor = None;
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

    fn focusable(&self) -> bool {
        !self.disabled
    }
}
