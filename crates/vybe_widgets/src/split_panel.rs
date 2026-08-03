//! SplitPanel — two children with a draggable splitter bar.
//!
//! Handles splitter drag internally. The IDE only needs to set the two
//! child panels and an initial split position.

use crate::layout::*;
use tiny_skia::*;

/// Two panels separated by a draggable splitter.
pub struct SplitPanel {
    pub id: WidgetId,
    rect: LayoutRect,
    horizontal: bool,
    split_pos: f32,
    splitter_width: f32,
    dragging: bool,
    min_size1: f32,
    min_size2: f32,
    panel1: Option<Box<dyn PanelWidget>>,
    panel2: Option<Box<dyn PanelWidget>>,
    splitter_color: (u8, u8, u8, u8),
    hover_color: (u8, u8, u8, u8),
    hovering: bool,
}

impl SplitPanel {
    /// Create a new SplitPanel.
    /// `horizontal = true` means left | right, `false` means top / bottom.
    pub fn new(horizontal: bool) -> Self {
        Self {
            id: WidgetId::next(),
            rect: LayoutRect::zero(),
            horizontal,
            split_pos: 200.0,
            splitter_width: 4.0,
            dragging: false,
            min_size1: 50.0,
            min_size2: 50.0,
            panel1: None,
            panel2: None,
            splitter_color: (51, 51, 51, 255),
            hover_color: (0, 122, 204, 255),
            hovering: false,
        }
    }

    pub fn set_panel1(&mut self, widget: Box<dyn PanelWidget>) {
        self.panel1 = Some(widget);
        self.relayout();
    }

    pub fn set_panel2(&mut self, widget: Box<dyn PanelWidget>) {
        self.panel2 = Some(widget);
        self.relayout();
    }

    pub fn set_split_pos(&mut self, pos: f32) {
        self.split_pos = pos;
        self.relayout();
    }

    pub fn split_pos(&self) -> f32 {
        self.split_pos
    }
    pub fn is_dragging(&self) -> bool {
        self.dragging
    }
    pub fn is_hovering(&self) -> bool {
        self.hovering
    }

    pub fn set_min_sizes(&mut self, min1: f32, min2: f32) {
        self.min_size1 = min1;
        self.min_size2 = min2;
    }

    pub fn set_splitter_width(&mut self, w: f32) {
        self.splitter_width = w;
    }

    pub fn panel1_ref(&self) -> Option<&dyn PanelWidget> {
        self.panel1.as_deref()
    }
    pub fn panel1_mut(&mut self) -> Option<&mut (dyn PanelWidget + 'static)> {
        self.panel1.as_deref_mut()
    }
    pub fn panel2_ref(&self) -> Option<&dyn PanelWidget> {
        self.panel2.as_deref()
    }
    pub fn panel2_mut(&mut self) -> Option<&mut (dyn PanelWidget + 'static)> {
        self.panel2.as_deref_mut()
    }

    fn splitter_rect(&self) -> LayoutRect {
        let sw = self.splitter_width;
        if self.horizontal {
            LayoutRect::new(self.rect.x + self.split_pos, self.rect.y, sw, self.rect.h)
        } else {
            LayoutRect::new(self.rect.x, self.rect.y + self.split_pos, self.rect.w, sw)
        }
    }

    fn relayout(&mut self) {
        let sw = self.splitter_width;
        if self.horizontal {
            let max_pos = (self.rect.w - sw - self.min_size2).max(self.min_size1);
            self.split_pos = self.split_pos.clamp(self.min_size1, max_pos);
            if let Some(p1) = &mut self.panel1 {
                p1.set_rect(LayoutRect::new(
                    self.rect.x,
                    self.rect.y,
                    self.split_pos,
                    self.rect.h,
                ));
            }
            if let Some(p2) = &mut self.panel2 {
                let x2 = self.rect.x + self.split_pos + sw;
                let w2 = (self.rect.w - self.split_pos - sw).max(0.0);
                p2.set_rect(LayoutRect::new(x2, self.rect.y, w2, self.rect.h));
            }
        } else {
            let max_pos = (self.rect.h - sw - self.min_size2).max(self.min_size1);
            self.split_pos = self.split_pos.clamp(self.min_size1, max_pos);
            if let Some(p1) = &mut self.panel1 {
                p1.set_rect(LayoutRect::new(
                    self.rect.x,
                    self.rect.y,
                    self.rect.w,
                    self.split_pos,
                ));
            }
            if let Some(p2) = &mut self.panel2 {
                let y2 = self.rect.y + self.split_pos + sw;
                let h2 = (self.rect.h - self.split_pos - sw).max(0.0);
                p2.set_rect(LayoutRect::new(self.rect.x, y2, self.rect.w, h2));
            }
        }
    }
}

impl PanelWidget for SplitPanel {
    
    fn find_rect(&self, name: &str) -> Option<LayoutRect> {
        if self.name() == name { return Some(self.rect()); }
        if let Some(w) = &self.panel1 {
            if let Some(r) = w.find_rect(name) { return Some(r); }
        }
        if let Some(w) = &self.panel2 {
            if let Some(r) = w.find_rect(name) { return Some(r); }
        }
        None
    }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.relayout();
    }

    fn rect(&self) -> LayoutRect {
        self.rect
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        if let Some(p1) = &mut self.panel1 {
            p1.render(ctx);
        }
        if let Some(p2) = &mut self.panel2 {
            p2.render(ctx);
        }
        // Splitter bar
        let sr = self.splitter_rect();
        let s = ctx.scale;
        let mut paint = Paint::default();
        let c = if self.hovering || self.dragging {
            self.hover_color
        } else {
            self.splitter_color
        };
        paint.set_color_rgba8(c.0, c.1, c.2, c.3);
        if let Some(rect) = Rect::from_xywh(sr.x * s, sr.y * s, sr.w * s, sr.h * s) {
            ctx.pixmap
                .fill_rect(rect, &paint, Transform::identity(), None);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let sr = self.splitter_rect();
        match event.kind {
            MouseEventKind::Press(MouseButton::Left) => {
                if sr.contains(event.x, event.y) {
                    self.dragging = true;
                    return true;
                }
            }
            MouseEventKind::Release(MouseButton::Left) => {
                if self.dragging {
                    self.dragging = false;
                    return true;
                }
            }
            MouseEventKind::Move => {
                if self.dragging {
                    if self.horizontal {
                        self.split_pos = event.x - self.rect.x;
                    } else {
                        self.split_pos = event.y - self.rect.y;
                    }
                    self.relayout();
                    return true;
                }
                self.hovering = sr.contains(event.x, event.y);
            }
            MouseEventKind::Scroll(_) => {}
            _ => {}
        }
        // Route to children
        if let Some(p1) = &mut self.panel1 {
            if p1.rect().contains(event.x, event.y) && p1.handle_mouse(event) {
                return true;
            }
        }
        if let Some(p2) = &mut self.panel2 {
            if p2.rect().contains(event.x, event.y) && p2.handle_mouse(event) {
                return true;
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if let Some(p1) = &mut self.panel1 {
            if p1.handle_key(event) {
                return true;
            }
        }
        if let Some(p2) = &mut self.panel2 {
            if p2.handle_key(event) {
                return true;
            }
        }
        false
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        if let Some(p1) = &mut self.panel1 {
            if p1.rect().contains(x, y) && p1.handle_scroll(delta, x, y) {
                return true;
            }
        }
        if let Some(p2) = &mut self.panel2 {
            if p2.rect().contains(x, y) && p2.handle_scroll(delta, x, y) {
                return true;
            }
        }
        false
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        let sr = self.splitter_rect();
        if sr.contains(x, y) || self.dragging {
            return if self.horizontal {
                winit::window::CursorIcon::ColResize
            } else {
                winit::window::CursorIcon::RowResize
            };
        }
        if let Some(p1) = &self.panel1 {
            if p1.rect().contains(x, y) {
                return p1.cursor_at(x, y);
            }
        }
        if let Some(p2) = &self.panel2 {
            if p2.rect().contains(x, y) {
                return p2.cursor_at(x, y);
            }
        }
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = Vec::new();
        if let Some(p1) = &mut self.panel1 {
            events.extend(p1.drain_events());
        }
        if let Some(p2) = &mut self.panel2 {
            events.extend(p2.drain_events());
        }
        events
    }
}
