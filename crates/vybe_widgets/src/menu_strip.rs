//! MenuStrip widget — horizontal menu bar with item areas.

use super::WidgetColors;
use super::layout::{
    KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent, MouseEventKind,
    PanelWidget, RenderContext, WidgetEvent, WidgetId };
use tiny_skia::*;

pub struct MenuStrip {
    pub items: Vec<String>,
    /// Estimated widths for each item (caller sets from text measurement).
    pub item_widths: Vec<f32>,
    pub active_index: Option<usize>,
    pub hover_index: Option<usize>,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent> }

impl MenuStrip {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            item_widths: Vec::new(),
            active_index: None,
            hover_index: None,
            width: 400.0,
            height: 24.0,
            colors: WidgetColors {
                background: (240, 240, 240, 255),
                ..WidgetColors::default()
            },
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new() }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Default item width when item_widths is not set.
    fn default_item_width(&self) -> f32 {
        60.0
    }

    /// Get width for a specific item.
    fn get_item_width(&self, index: usize) -> f32 {
        if index < self.item_widths.len() {
            self.item_widths[index]
        } else {
            self.default_item_width()
        }
    }

    /// X position for a specific item.
    pub fn item_x(&self, index: usize) -> f32 {
        let mut x = 0.0;
        for i in 0..index {
            x += self.get_item_width(i);
        }
        x
    }

    /// Paint — menu bar background, item hover/active highlight, bottom border.
    /// Item text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Item highlights
        for i in 0..self.items.len() {
            let ix = x + self.item_x(i);
            let iw = self.get_item_width(i);
            let is_active = self.active_index == Some(i);
            let is_hovered = self.hover_index == Some(i);

            if is_active {
                // Active item: accent background
                let (r, g, b, _) = self.colors.accent;
                paint.set_color_rgba8(r, g, b, 40);
                if let Some(rect) = Rect::from_xywh(ix, y, iw, self.height) {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
            } else if is_hovered {
                // Hover: subtle highlight
                paint.set_color_rgba8(0, 0, 0, 15);
                if let Some(rect) = Rect::from_xywh(ix, y, iw, self.height) {
                    pixmap.fill_rect(rect, &paint, ts, None);
                }
            }
        }

        // Bottom border
        paint.set_color_rgba8(210, 210, 210, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x + self.width, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Hit test — returns item index at position.
    pub fn hit_test(&self, mx: f32, _my: f32) -> Option<usize> {
        if _my < 0.0 || _my > self.height {
            return None;
        }
        let mut cx = 0.0;
        for i in 0..self.items.len() {
            let iw = self.get_item_width(i);
            if mx >= cx && mx < cx + iw {
                return Some(i);
            }
            cx += iw;
        }
        None
    }

    /// Update hover state on mouse move.
    pub fn mouse_move(&mut self, mx: f32, my: f32) {
        self.hover_index = self.hit_test(mx, my);
    }
}

impl PanelWidget for MenuStrip {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
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
            let ix = r.x + self.item_x(i) + 8.0;
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                item,
                ix,
                r.y + 4.0,
                12.0,
                col,
                ctx.scale,
            );
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        if !r.contains(event.x, event.y) {
            self.hover_index = None;
            return false;
        }
        let lx = event.x - r.x;
        let ly = event.y - r.y;
        match event.kind {
            MouseEventKind::Move => {
                self.mouse_move(lx, ly);
            }
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                if let Some(idx) = self.hit_test(lx, ly) {
                    self.active_index = Some(idx);
                    self.pending_events
                        .push(WidgetEvent::MenuItemClicked(self.name.clone(), idx));
                    return true;
                }
            }
            _ => {}
        }
        false
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }
    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }
}
