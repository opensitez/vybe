//! ListBox widget — standalone tiny-skia rendered list box.

use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent,
    MouseEventKind, PanelWidget, RenderContext, SelectionMode, WidgetCommand, WidgetEvent,
    WidgetId };
use super::{WidgetColors, rounded_rect_path};
use tiny_skia::*;

pub struct ListBox {
    pub items: Vec<String>,
    pub selected_index: Option<usize>,
    /// All selected indices (used for multi-select modes; always includes selected_index).
    pub selected_indices: Vec<usize>,
    pub selection_mode: SelectionMode,
    pub item_height: f32,
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
    /// Anchor index for Shift+click range selection.
    range_anchor: Option<usize> }

impl ListBox {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            selected_index: None,
            selected_indices: Vec::new(),
            selection_mode: SelectionMode::Single,
            item_height: 18.0,
            scroll_offset: 0.0,
            focused: false,
            hovered: false,
            width: 120.0,
            height: 120.0,
            colors: WidgetColors::default(),
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
            range_anchor: None }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }
    pub fn with_selection_mode(mut self, mode: SelectionMode) -> Self {
        self.selection_mode = mode;
        self
    }

    /// Is the given index currently selected?
    pub fn is_selected(&self, idx: usize) -> bool {
        match self.selection_mode {
            SelectionMode::Single => self.selected_index == Some(idx),
            _ => self.selected_indices.contains(&idx) }
    }

    /// Paint the listbox — white background, inset border, selection highlight.
    /// Item text is drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // White background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 1.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Selection highlight bars
        for i in 0..self.items.len() {
            if !self.is_selected(i) {
                continue;
            }
            let item_y = y + 1.0 + (i as f32 * self.item_height) - self.scroll_offset;
            let bar_top = item_y.max(y + 1.0);
            let bar_bottom = (item_y + self.item_height).min(y + self.height - 1.0);
            if bar_top < bar_bottom {
                let (r, g, b, _) = self.colors.accent;
                paint.set_color_rgba8(r, g, b, 60);
                if let Some(rect) =
                    Rect::from_xywh(x + 1.0, bar_top, self.width - 2.0, bar_bottom - bar_top)
                {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
                paint.set_color_rgba8(r, g, b, 160);
                let mut stroke = Stroke::default();
                stroke.width = 1.0;
                if let Some(rect_path) = rounded_rect_path(
                    x + 1.0,
                    bar_top,
                    self.width - 2.0,
                    bar_bottom - bar_top,
                    0.0,
                ) {
                    pixmap.stroke_path(&rect_path, &paint, &stroke, ts, None);
                }
            }
        }

        // Inset border (3D sunken effect: dark top-left, light bottom-right)
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        // Dark top and left edges (inset)
        paint.set_color_rgba8(130, 135, 144, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x, y);
        pb.line_to(x + self.width, y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Light bottom and right edges
        paint.set_color_rgba8(255, 255, 255, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Inner border
        paint.set_color_rgba8(160, 160, 160, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x + 1.0, y + self.height - 1.0);
        pb.line_to(x + 1.0, y + 1.0);
        pb.line_to(x + self.width - 1.0, y + 1.0);
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

    /// Hit-test: given local coordinates, return item index if within bounds.
    pub fn hit_test(&self, x: f32, y: f32) -> Option<usize> {
        if x < 0.0 || y < 0.0 || x > self.width || y > self.height {
            return None;
        }
        let adjusted_y = y - 1.0 + self.scroll_offset;
        if adjusted_y < 0.0 {
            return None;
        }
        let idx = (adjusted_y / self.item_height) as usize;
        if idx < self.items.len() {
            Some(idx)
        } else {
            None
        }
    }

    /// Handle click at (x, y) relative to widget origin. Returns item index if hit.
    pub fn click(&mut self, x: f32, y: f32) -> Option<usize> {
        if let Some(idx) = self.hit_test(x, y) {
            self.select_single(idx);
            Some(idx)
        } else {
            None
        }
    }

    /// Select a single index (Single mode logic).
    fn select_single(&mut self, idx: usize) {
        self.selected_index = Some(idx);
        self.selected_indices = vec![idx];
        self.range_anchor = Some(idx);
    }

    /// Toggle selection of an index (for MultiSimple / Ctrl+click in MultiExtended).
    fn toggle_index(&mut self, idx: usize) {
        if let Some(pos) = self.selected_indices.iter().position(|&i| i == idx) {
            self.selected_indices.remove(pos);
        } else {
            self.selected_indices.push(idx);
        }
        self.selected_indices.sort();
        self.selected_index = self.selected_indices.first().copied();
        self.range_anchor = Some(idx);
    }

    /// Select a contiguous range from the anchor to idx (for Shift+click in MultiExtended).
    fn select_range_to(&mut self, idx: usize) {
        let anchor = self.range_anchor.unwrap_or(0);
        let (start, end) = if anchor <= idx {
            (anchor, idx)
        } else {
            (idx, anchor)
        };
        self.selected_indices = (start..=end).collect();
        self.selected_index = Some(idx);
    }

    /// Y position of an item relative to widget origin (for text placement).
    pub fn item_y(&self, index: usize) -> f32 {
        1.0 + (index as f32 * self.item_height) - self.scroll_offset
    }
}

impl PanelWidget for ListBox {
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
        // Draw item text
        let (fr, fg, fb, _) = self.colors.foreground;
        let col = cosmic_text::Color::rgba(fr, fg, fb, 255);
        for (i, item) in self.items.iter().enumerate() {
            let iy = r.y + self.item_y(i);
            if iy + self.item_height < r.y || iy > r.y + r.h {
                continue;
            }
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                item,
                r.x + 4.0,
                iy + 1.0,
                12.0,
                col,
                ctx.scale,
            );
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        if !r.contains(event.x, event.y) {
            return false;
        }
        let lx = event.x - r.x;
        let ly = event.y - r.y;
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            if let Some(idx) = self.hit_test(lx, ly) {
                match self.selection_mode {
                    SelectionMode::Single => {
                        self.select_single(idx);
                    }
                    SelectionMode::MultiSimple => {
                        self.toggle_index(idx);
                    }
                    SelectionMode::MultiExtended => {
                        if event.shift && self.range_anchor.is_some() {
                            self.select_range_to(idx);
                        } else if event.cmd {
                            self.toggle_index(idx);
                        } else {
                            self.select_single(idx);
                        }
                    }
                }
                self.pending_events
                    .push(WidgetEvent::ListBoxSelected(self.name.clone(), idx));
                return true;
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused {
            return false;
        }
        use winit::event::ElementState;
        use winit::keyboard::{Key, NamedKey};
        if event.state != ElementState::Pressed {
            return false;
        }
        match &event.logical_key {
            Key::Named(NamedKey::ArrowDown) => {
                let next = self
                    .selected_index
                    .map(|i| (i + 1).min(self.items.len().saturating_sub(1)))
                    .unwrap_or(0);
                if next < self.items.len() {
                    match self.selection_mode {
                        SelectionMode::Single => self.select_single(next),
                        SelectionMode::MultiSimple => self.select_single(next),
                        SelectionMode::MultiExtended => {
                            if event.shift {
                                if self.range_anchor.is_none() {
                                    self.range_anchor = self.selected_index;
                                }
                                self.select_range_to(next);
                            } else {
                                self.select_single(next);
                            }
                        }
                    }
                    self.pending_events
                        .push(WidgetEvent::ListBoxSelected(self.name.clone(), next));
                }
                true
            }
            Key::Named(NamedKey::ArrowUp) => {
                let prev = self
                    .selected_index
                    .map(|i| i.saturating_sub(1))
                    .unwrap_or(0);
                if prev < self.items.len() {
                    match self.selection_mode {
                        SelectionMode::Single => self.select_single(prev),
                        SelectionMode::MultiSimple => self.select_single(prev),
                        SelectionMode::MultiExtended => {
                            if event.shift {
                                if self.range_anchor.is_none() {
                                    self.range_anchor = self.selected_index;
                                }
                                self.select_range_to(prev);
                            } else {
                                self.select_single(prev);
                            }
                        }
                    }
                    self.pending_events
                        .push(WidgetEvent::ListBoxSelected(self.name.clone(), prev));
                }
                true
            }
            _ => false }
    }

    fn handle_scroll(&mut self, _x: f32, _y: f32, delta_y: f32) -> bool {
        self.scroll_offset = (self.scroll_offset - delta_y * 20.0).max(0.0);
        let max_scroll = (self.items.len() as f32 * self.item_height - self.height).max(0.0);
        self.scroll_offset = self.scroll_offset.min(max_scroll);
        true
    }

    fn focusable(&self) -> bool {
        true
    }
    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetSelectedIndex(i) => {
                if *i < self.items.len() && self.selected_index != Some(*i) {
                    self.select_single(*i);
                    self.pending_events
                        .push(WidgetEvent::ListBoxSelected(self.name.clone(), *i));
                }
                CommandValue::None
            }
            WidgetCommand::GetValue => CommandValue::Index(self.selected_index.unwrap_or(0)),
            WidgetCommand::AddItem(s) => {
                self.items.push(s.clone());
                CommandValue::None
            }
            WidgetCommand::RemoveItem(i) => {
                if *i < self.items.len() {
                    self.items.remove(*i);
                    self.selected_indices.retain(|&idx| idx != *i);
                    // Adjust indices above the removed item
                    for idx in &mut self.selected_indices {
                        if *idx > *i {
                            *idx -= 1;
                        }
                    }
                    if self.selected_index == Some(*i) {
                        self.selected_index = self.selected_indices.first().copied();
                    } else if let Some(ref mut si) = self.selected_index {
                        if *si > *i {
                            *si -= 1;
                        }
                    }
                }
                CommandValue::None
            }
            WidgetCommand::ClearItems => {
                self.items.clear();
                self.selected_index = None;
                self.selected_indices.clear();
                self.range_anchor = None;
                CommandValue::None
            }
            WidgetCommand::GetText => {
                let t = self
                    .selected_index
                    .and_then(|i| self.items.get(i))
                    .cloned()
                    .unwrap_or_default();
                CommandValue::Text(t)
            }
            WidgetCommand::Custom(key, _val) => match key.as_str() {
                "GetSelectedIndices" => {
                    let s = self
                        .selected_indices
                        .iter()
                        .map(|i| i.to_string())
                        .collect::<Vec<_>>()
                        .join(",");
                    CommandValue::Text(s)
                }
                _ => CommandValue::None },
            _ => CommandValue::None }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }
}
