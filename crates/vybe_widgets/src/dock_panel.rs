//! DockPanel — arranges children by dock position (Left, Right, Top, Bottom, Fill).
//!
//! Children are processed in order. Each non-Fill child takes space from the
//! remaining rectangle. The last Fill child gets whatever remains.

use crate::layout::*;
use tiny_skia::*;

/// A docked child with its position and requested size.
pub struct DockChild {
    pub dock: Dock,
    pub size: f32,
    pub min_size: f32,
    pub visible: bool,
    pub widget: Box<dyn PanelWidget>,
}

/// Container that arranges children by dock position (like WPF DockPanel).
pub struct DockPanel {
    pub id: WidgetId,
    rect: LayoutRect,
    children: Vec<DockChild>,
    bg_color: (u8, u8, u8, u8),
}

impl DockPanel {
    pub fn new() -> Self {
        Self {
            id: WidgetId::next(),
            rect: LayoutRect::zero(),
            children: Vec::new(),
            bg_color: (30, 30, 30, 255),
        }
    }

    pub fn set_background(&mut self, r: u8, g: u8, b: u8, a: u8) {
        self.bg_color = (r, g, b, a);
    }

    pub fn add(&mut self, dock: Dock, size: f32, widget: Box<dyn PanelWidget>) {
        self.children.push(DockChild {
            dock,
            size,
            min_size: 0.0,
            visible: true,
            widget,
        });
        self.relayout();
    }

    pub fn add_with_min(
        &mut self,
        dock: Dock,
        size: f32,
        min_size: f32,
        widget: Box<dyn PanelWidget>,
    ) {
        self.children.push(DockChild {
            dock,
            size,
            min_size,
            visible: true,
            widget,
        });
        self.relayout();
    }

    pub fn child(&self, index: usize) -> &dyn PanelWidget {
        &*self.children[index].widget
    }

    pub fn child_mut(&mut self, index: usize) -> &mut dyn PanelWidget {
        &mut *self.children[index].widget
    }

    pub fn child_count(&self) -> usize {
        self.children.len()
    }

    /// Set the dock size of a child. Triggers relayout.
    pub fn set_child_size(&mut self, index: usize, size: f32) {
        if index < self.children.len() {
            let min = self.children[index].min_size;
            self.children[index].size = size.max(min);
            self.relayout();
        }
    }

    /// Show or hide a child. Hidden children take no space.
    pub fn set_child_visible(&mut self, index: usize, visible: bool) {
        if index < self.children.len() {
            self.children[index].visible = visible;
            self.relayout();
        }
    }

    pub fn is_child_visible(&self, index: usize) -> bool {
        self.children.get(index).map_or(false, |c| c.visible)
    }

    fn relayout(&mut self) {
        let mut remaining = self.rect;
        for child in &mut self.children {
            if !child.visible {
                child.widget.set_rect(LayoutRect::zero());
                continue;
            }
            let child_rect = match child.dock {
                Dock::Left => remaining.take_left(child.size),
                Dock::Right => remaining.take_right(child.size),
                Dock::Top => remaining.take_top(child.size),
                Dock::Bottom => remaining.take_bottom(child.size),
                Dock::Fill => {
                    let r = remaining;
                    remaining = LayoutRect::zero();
                    r
                }
            };
            child.widget.set_rect(child_rect);
        }
    }
}

impl PanelWidget for DockPanel {
    
    fn find_rect(&self, name: &str) -> Option<LayoutRect> {
        if self.name() == name { return Some(self.rect()); }
        for child in &self.children {
            if let Some(r) = child.widget.find_rect(name) { return Some(r); }
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
        // Background
        let (r, g, b, a) = self.bg_color;
        let s = ctx.scale;
        let mut paint = Paint::default();
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(
            self.rect.x * s,
            self.rect.y * s,
            self.rect.w * s,
            self.rect.h * s,
        ) {
            ctx.pixmap
                .fill_rect(rect, &paint, Transform::identity(), None);
        }
        // Children
        for child in &mut self.children {
            if child.visible {
                child.widget.render(ctx);
            }
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        // Route to children in reverse order (last = topmost)
        for child in self.children.iter_mut().rev() {
            if child.visible && child.widget.rect().contains(event.x, event.y) {
                if child.widget.handle_mouse(event) {
                    return true;
                }
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        for child in &mut self.children {
            if child.visible && child.widget.handle_key(event) {
                return true;
            }
        }
        false
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        for child in self.children.iter_mut().rev() {
            if child.visible && child.widget.rect().contains(x, y) {
                if child.widget.handle_scroll(delta, x, y) {
                    return true;
                }
            }
        }
        false
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        for child in self.children.iter().rev() {
            if child.visible && child.widget.rect().contains(x, y) {
                return child.widget.cursor_at(x, y);
            }
        }
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = Vec::new();
        for child in &mut self.children {
            if child.visible {
                events.extend(child.widget.drain_events());
            }
        }
        events
    }
}
