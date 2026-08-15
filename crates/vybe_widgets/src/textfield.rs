//! Text input widget — standalone tiny-skia rendered text field.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, TextAlign, WidgetCommand, WidgetEvent, WidgetId,
};
use cosmic_text::Color as CosmicColor;

pub struct TextInput {
    pub value: String,
    pub placeholder: String,
    pub cursor: usize,
    /// Selection anchor (the other end of the selection range). `None` = no selection.
    pub selection_anchor: Option<usize>,
    pub password: bool,
    /// Whether a newline is content rather than a terminator. A memo, a
    /// `<textarea>` and a WinForms `RichTextBox` are all this control with the
    /// flag set: the text is top-aligned, `Enter` inserts, and the caret,
    /// selection and hit-test work in (line, column) instead of one offset.
    pub multiline: bool,
    pub disabled: bool,
    pub read_only: bool,
    pub max_length: Option<usize>,
    pub focused: bool,
    pub hovered: bool,
    pub colors: WidgetColors,
    /// The field's text style — same reasoning as `Button::font`.
    pub font: crate::ide_text::FontSpec,
    /// CSS `text-align` — VCL's `Alignment`, WinForms' `TextAlign`. A
    /// calculator display is the canonical case and the reason it is here
    /// rather than on the label alone.
    pub text_align: TextAlign,
    pub width: f32,
    pub height: f32,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
    /// Whether a mouse drag selection is in progress.
    dragging: bool,
    /// Where each line was actually DRAWN, as an offset from the field's left
    /// edge, recorded by the last render.
    ///
    /// The caret, the selection rect and the hit test all have to land on the
    /// same pixels as the glyphs, and with an alignment that offset is no
    /// longer the constant `padding` — it depends on the measured width of the
    /// line, which needs a font system. Render has one; the event path does
    /// not. So the number is measured once, where it is knowable, and read
    /// back where it is needed, rather than estimated twice and disagreed on.
    line_origins: Vec<f32>,
}

impl TextInput {
    pub fn new() -> Self {
        Self {
            value: String::new(),
            placeholder: String::new(),
            cursor: 0,
            selection_anchor: None,
            password: false,
            multiline: false,
            disabled: false,
            read_only: false,
            max_length: None,
            focused: false,
            hovered: false,
            colors: WidgetColors::default(),
            font: crate::ide_text::FontSpec::sans(14.0),
            text_align: TextAlign::Left,
            width: 200.0,
            height: 24.0,
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
            dragging: false,
            line_origins: Vec::new(),
        }
    }

    /// Where a line of `width` pixels starts, as an offset from the left edge.
    ///
    /// The one place the alignment is turned into a number. `Left` is the
    /// padding a field has always had; the other two measure from the opposite
    /// edge, which is what makes a right-aligned display line up on the right
    /// however long the number is.
    fn line_origin(&self, line_width: f32) -> f32 {
        const PADDING: f32 = 4.0;
        match self.text_align {
            TextAlign::Left => PADDING,
            TextAlign::Center => ((self.rect.w - line_width) / 2.0).max(PADDING),
            TextAlign::Right => (self.rect.w - PADDING - line_width).max(PADDING),
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

    /// The baseline-to-baseline step the text renderer actually uses, so the
    /// caret and the hit-test land on the same rows the glyphs do.
    /// Mirrors `Metrics::new(size, size * 1.3)` in `ide_text`.
    fn line_height(&self) -> f32 {
        self.font.resolved_line_height()
    }

    /// `(start_byte, text)` per line. A single-line field is simply one line,
    /// so everything below is written once for both cases rather than branched.
    ///
    /// A trailing `\r` is dropped from the TEXT but still counted in the
    /// offsets: `\r\n` is how Delphi and Windows spell a line break, and it is
    /// terminator, not content — drawing it would put a box at every EOL.
    fn lines(&self) -> Vec<(usize, &str)> {
        let mut out = Vec::new();
        let mut start = 0usize;
        for raw in self.value.split('\n') {
            out.push((start, raw.strip_suffix('\r').unwrap_or(raw)));
            start += raw.len() + 1; // + the '\n' itself
        }
        out
    }

    /// The cursor as `(line index, byte offset within that line)`.
    fn cursor_line_col(&self) -> (usize, usize) {
        let cursor = self.cursor.min(self.value.len());
        let before = &self.value[..cursor];
        let line = before.matches('\n').count();
        let col = before.rfind('\n').map_or(cursor, |nl| cursor - nl - 1);
        (line, col)
    }

    /// Move the caret a line at a time, keeping its column where it can.
    /// Single-line fields have nowhere to go, so this is a no-op for them.
    fn move_vertical(&mut self, delta: isize) {
        let target = {
            let lines = self.lines();
            let (line, col) = self.cursor_line_col();
            let want = line as isize + delta;
            if want < 0 || want >= lines.len() as isize {
                None
            } else {
                let (start, text) = lines[want as usize];
                let mut col = col.min(text.len());
                while col > 0 && !text.is_char_boundary(col) {
                    col -= 1;
                }
                Some(start + col)
            }
        };
        if let Some(cursor) = target {
            self.cursor = cursor;
        }
    }

    /// Byte offset for a point in the widget's own coordinates.
    ///
    /// The column is the same average-advance estimate this control has always
    /// used — it has no font system on the event path — but it is now applied
    /// to the line the point falls on instead of to the whole value.
    fn offset_at(&self, x: f32, y: f32) -> usize {
        let padding = 4.0;
        let lines = self.lines();
        let index = if self.multiline {
            let local_y = y - self.rect.y - padding;
            ((local_y / self.line_height()).floor().max(0.0) as usize).min(lines.len() - 1)
        } else {
            0
        };
        let (start, text) = lines[index];
        if text.is_empty() {
            return start;
        }
        // Measured at render, because that is where a font system exists —
        // see `line_origins`. A field that has never been drawn falls back to
        // the padding, which is the left-aligned answer and the old constant.
        let origin = self.line_origins.get(index).copied().unwrap_or(padding);
        let local_x = x - self.rect.x - origin;
        let chars = text.chars().count();
        let advance = self.width / chars.max(1) as f32;
        let column = ((local_x / advance).round().max(0.0) as usize).min(chars);
        start
            + text
                .char_indices()
                .nth(column)
                .map_or(text.len(), |(i, _)| i)
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
            let w =
                super::ide_text::measure_text_spec(font_system, &text[..pos], &self.font, scale);
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
        let is_ph = self.is_placeholder();
        let (cr, cg, cb, _) = if is_ph {
            self.colors.placeholder
        } else {
            self.colors.foreground
        };
        // A memo fills from the top; a one-line field centres in its box.
        let ty = if self.multiline {
            r.y + padding
        } else {
            r.y + (r.h - self.font.size) / 2.0 - 1.0
        };
        let line_h = self.line_height();
        // One line for a single-line field, so the loops below are the general
        // case rather than a second implementation.
        let lines: Vec<(usize, String)> = if is_ph {
            vec![(0, self.placeholder.clone())]
        } else {
            self.lines()
                .into_iter()
                .map(|(start, text)| {
                    let shown = if self.password {
                        "\u{2022}".repeat(text.chars().count())
                    } else {
                        text.to_string()
                    };
                    (start, shown)
                })
                .collect()
        };
        // Width of the first `bytes` of a line as drawn.
        let prefix_x = |ctx: &mut RenderContext, text: &str, bytes: usize| -> f32 {
            let end = bytes.min(text.len());
            if end == 0 {
                return 0.0;
            }
            super::ide_text::measure_text_spec(ctx.font_system, &text[..end], &self.font, ctx.scale)
        };
        // Each line's own origin, from its own measured width — a shorter line
        // in a right-aligned field starts further right. Recorded on the field
        // so the event path can reach the same number.
        let origins: Vec<f32> = lines
            .iter()
            .map(|(_, text)| {
                let width = prefix_x(ctx, text, text.len());
                self.line_origin(width)
            })
            .collect();
        let origin_of = |i: usize| origins.get(i).copied().unwrap_or(padding);

        // Selection highlight — one rect per line it covers.
        if self.focused && !is_ph {
            if let Some((sel_start, sel_end)) = self.selection_range() {
                for (i, (start, text)) in lines.iter().enumerate() {
                    let (from, to) = (sel_start.max(*start), sel_end.min(start + text.len()));
                    if from >= to {
                        continue;
                    }
                    let x_start = prefix_x(ctx, text, from - start);
                    let x_end = prefix_x(ctx, text, to - start);
                    let sx = (r.x + origin_of(i) + x_start) * ctx.scale;
                    let sw = (x_end - x_start) * ctx.scale;
                    let (sy, sh) = if self.multiline {
                        ((ty + i as f32 * line_h) * ctx.scale, line_h * ctx.scale)
                    } else {
                        ((r.y + 2.0) * ctx.scale, (r.h - 4.0) * ctx.scale)
                    };
                    if let Some(rect) = tiny_skia::Rect::from_xywh(sx, sy, sw.max(1.0), sh) {
                        let mut paint = tiny_skia::Paint::default();
                        paint.set_color_rgba8(51, 153, 255, 100); // blue selection highlight
                        ctx.pixmap
                            .fill_rect(rect, &paint, tiny_skia::Transform::identity(), None);
                    }
                }
            }
        }

        for (i, (_, text)) in lines.iter().enumerate() {
            let y = ty + i as f32 * line_h;
            // Stop at the bottom edge instead of painting outside the control.
            if self.multiline && y + line_h > r.y + r.h {
                break;
            }
            super::ide_text::draw_text_spec(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                text,
                r.x + origin_of(i),
                y,
                &self.font,
                CosmicColor::rgba(cr, cg, cb, 255),
                ctx.scale,
            );
        }

        // Cursor
        if self.focused {
            let (line, col) = self.cursor_line_col();
            let cursor_x = lines
                .get(line)
                .map_or(0.0, |(_, text)| prefix_x(ctx, text, col));
            let cx = (r.x + origin_of(line) + cursor_x) * ctx.scale;
            let (cy_top, cy_bot) = if self.multiline {
                let top = ty + line as f32 * line_h;
                ((top * ctx.scale), (top + line_h) * ctx.scale)
            } else {
                ((r.y + 3.0) * ctx.scale, (r.y + r.h - 3.0) * ctx.scale)
            };
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
        // Hand the measured origins to the event path.
        self.line_origins = origins;
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            self.dragging = false;
            return false;
        }
        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                self.focused = true;
                self.cursor = self.offset_at(event.x, event.y);
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
                    if self.selection_anchor.is_none() {
                        self.selection_anchor = Some(self.cursor);
                    }
                    self.cursor = self.offset_at(event.x, event.y);
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
            // In a memo a newline is CONTENT. A single-line field keeps
            // ignoring Enter, which is what submits a form.
            Key::Named(NamedKey::Enter) if self.multiline => {
                self.insert("\n");
                self.pending_events.push(WidgetEvent::TextChanged(
                    self.name.clone(),
                    self.value.clone(),
                ));
                true
            }
            Key::Named(NamedKey::ArrowUp) | Key::Named(NamedKey::ArrowDown) if self.multiline => {
                if is_shift {
                    if self.selection_anchor.is_none() {
                        self.selection_anchor = Some(self.cursor);
                    }
                } else {
                    self.clear_selection();
                }
                self.move_vertical(
                    if matches!(&event.key_without_modifiers, Key::Named(NamedKey::ArrowUp)) {
                        -1
                    } else {
                        1
                    },
                );
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
                // Home is the start of the LINE in a memo, of the value in a field.
                self.cursor = if self.multiline {
                    self.cursor - self.cursor_line_col().1
                } else {
                    0
                };
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
                self.cursor = if self.multiline {
                    let (line, _) = self.cursor_line_col();
                    let lines = self.lines();
                    lines[line].0 + lines[line].1.len()
                } else {
                    self.value.len()
                };
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
            WidgetCommand::Custom(key, val) if self.font.apply_command(key, val) => {
                CommandValue::None
            }
            WidgetCommand::Custom(key, CommandValue::Text(value)) if key == "SetTextAlign" => {
                if let Some(align) = TextAlign::from_css(value) {
                    self.text_align = align;
                }
                CommandValue::None
            }
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
