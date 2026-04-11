//! ListView widget — column-based list with header row.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent, WidgetId, WidgetCommand, CommandValue};

pub struct ListView {
    pub items: Vec<String>,
    pub columns: Vec<String>,
    pub selected_index: Option<usize>,
    pub item_height: f32,
    pub header_height: f32,
    pub scroll_offset: f32,
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

impl ListView {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            columns: Vec::new(),
            selected_index: None,
            item_height: 20.0,
            header_height: 24.0,
            scroll_offset: 0.0,
            focused: false,
            hovered: false,
            width: 200.0,
            height: 150.0,
            colors: WidgetColors::default(),
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Paint the list view — white background, header, column dividers, selection.
    /// Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // White background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Header background (darker gray)
        paint.set_color_rgba8(230, 230, 230, 255);
        if let Some(rect) = Rect::from_xywh(x + 1.0, y + 1.0, self.width - 2.0, self.header_height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Header bottom border
        paint.set_color_rgba8(180, 180, 180, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.header_height);
        pb.line_to(x + self.width, y + self.header_height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Column dividers
        if !self.columns.is_empty() {
            let col_w = self.width / self.columns.len() as f32;
            paint.set_color_rgba8(200, 200, 200, 255);
            for i in 1..self.columns.len() {
                let cx = x + i as f32 * col_w;
                let mut pb = PathBuilder::new();
                pb.move_to(cx, y);
                pb.line_to(cx, y + self.height);
                if let Some(path) = pb.finish() {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
            }
        }

        // Selection highlight
        if let Some(idx) = self.selected_index {
            let item_y = y + self.header_height + 1.0 + (idx as f32 * self.item_height) - self.scroll_offset;
            let bar_top = item_y.max(y + self.header_height + 1.0);
            let bar_bottom = (item_y + self.item_height).min(y + self.height - 1.0);
            if bar_top < bar_bottom {
                let (r, g, b, _) = self.colors.accent;
                paint.set_color_rgba8(r, g, b, 50);
                if let Some(rect) = Rect::from_xywh(x + 1.0, bar_top, self.width - 2.0, bar_bottom - bar_top) {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
            }
        }

        // Inset border (sunken)
        paint.set_color_rgba8(130, 135, 144, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x, y);
        pb.line_to(x + self.width, y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
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
            if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 1.0) {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Handle click — returns item index.
    pub fn click(&mut self, x: f32, y: f32) -> Option<usize> {
        if x < 0.0 || y < 0.0 || x > self.width || y > self.height {
            return None;
        }
        let adjusted_y = y - self.header_height - 1.0 + self.scroll_offset;
        if adjusted_y < 0.0 {
            return None;
        }
        let idx = (adjusted_y / self.item_height) as usize;
        if idx < self.items.len() {
            self.selected_index = Some(idx);
            Some(idx)
        } else {
            None
        }
    }

    /// Column width for layout.
    pub fn column_width(&self) -> f32 {
        if self.columns.is_empty() {
            self.width
        } else {
            self.width / self.columns.len() as f32
        }
    }
}

impl PanelWidget for ListView {
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
        // Draw column headers
        let (fr, fg, fb, _) = self.colors.foreground;
        let col = cosmic_text::Color::rgba(fr, fg, fb, 255);
        let cw = self.column_width();
        for (i, header) in self.columns.iter().enumerate() {
            super::ide_text::draw_text(ctx.pixmap, ctx.font_system, ctx.swash_cache, header, r.x + i as f32 * cw + 4.0, r.y + 4.0, 12.0, col, ctx.scale);
        }
        // Draw items
        for (i, item) in self.items.iter().enumerate() {
            let iy = r.y + self.header_height + 1.0 + i as f32 * self.item_height - self.scroll_offset;
            if iy + self.item_height < r.y + self.header_height || iy > r.y + r.h { continue; }
            super::ide_text::draw_text(ctx.pixmap, ctx.font_system, ctx.swash_cache, item, r.x + 4.0, iy + 2.0, 12.0, col, ctx.scale);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        if !r.contains(event.x, event.y) { return false; }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            if let Some(idx) = self.click(event.x - r.x, event.y - r.y) {
                self.pending_events.push(WidgetEvent::ListViewSelected(self.name.clone(), idx));
                return true;
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused { return false; }
        use winit::keyboard::{Key, NamedKey};
        match &event.logical_key {
            Key::Named(NamedKey::ArrowDown) => {
                let next = self.selected_index.map(|i| (i + 1).min(self.items.len().saturating_sub(1))).unwrap_or(0);
                if next < self.items.len() {
                    self.selected_index = Some(next);
                    self.pending_events.push(WidgetEvent::ListViewSelected(self.name.clone(), next));
                }
                true
            }
            Key::Named(NamedKey::ArrowUp) => {
                let prev = self.selected_index.map(|i| i.saturating_sub(1)).unwrap_or(0);
                if prev < self.items.len() {
                    self.selected_index = Some(prev);
                    self.pending_events.push(WidgetEvent::ListViewSelected(self.name.clone(), prev));
                }
                true
            }
            _ => false,
        }
    }

    fn focusable(&self) -> bool { true }
    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetSelectedIndex(i) => {
                if *i < self.items.len() && self.selected_index != Some(*i) {
                    self.selected_index = Some(*i);
                    self.pending_events.push(WidgetEvent::ListViewSelected(
                        self.name.clone(),
                        *i,
                    ));
                }
                CommandValue::None
            }
            WidgetCommand::GetValue => CommandValue::Index(self.selected_index.unwrap_or(0)),
            WidgetCommand::AddItem(s) => { self.items.push(s.clone()); CommandValue::None }
            WidgetCommand::RemoveItem(i) => { if *i < self.items.len() { self.items.remove(*i); } CommandValue::None }
            WidgetCommand::ClearItems => { self.items.clear(); self.selected_index = None; CommandValue::None }
            WidgetCommand::GetText => {
                let t = self.selected_index.and_then(|i| self.items.get(i)).cloned().unwrap_or_default();
                CommandValue::Text(t)
            }
            _ => CommandValue::None,
        }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> { std::mem::take(&mut self.pending_events) }
}
