//! WrapPanel — arranges children in sequential position, wrapping to the next
//! line when the edge of the panel is reached.
//!
//! Unlike FlowLayoutPanel (which preserves each child's size), WrapPanel can
//! optionally enforce uniform item width/height for a grid-like appearance.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetCommand,
    WidgetEvent, WidgetId,
};
use super::stack_panel::Orientation;
use tiny_skia::*;

pub struct WrapPanel {
    pub orientation: Orientation,
    /// Spacing between children in pixels.
    pub spacing: f32,
    /// Padding inside the panel edges.
    pub padding: f32,
    /// If set, all children use this width. Otherwise each child keeps its own width.
    pub item_width: Option<f32>,
    /// If set, all children use this height. Otherwise each child keeps its own height.
    pub item_height: Option<f32>,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    children: Vec<Box<dyn PanelWidget>>,
}

impl WrapPanel {
    pub fn new(orientation: Orientation) -> Self {
        Self {
            orientation,
            spacing: 4.0,
            padding: 4.0,
            item_width: None,
            item_height: None,
            colors: WidgetColors {
                background: (240, 240, 240, 0), // transparent by default
                ..WidgetColors::default()
            },
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            children: Vec::new(),
        }
    }

    pub fn horizontal() -> Self {
        Self::new(Orientation::Horizontal)
    }
    pub fn vertical() -> Self {
        Self::new(Orientation::Vertical)
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }
    pub fn with_spacing(mut self, spacing: f32) -> Self {
        self.spacing = spacing;
        self
    }
    pub fn with_padding(mut self, padding: f32) -> Self {
        self.padding = padding;
        self
    }
    pub fn with_item_width(mut self, w: f32) -> Self {
        self.item_width = Some(w);
        self
    }
    pub fn with_item_height(mut self, h: f32) -> Self {
        self.item_height = Some(h);
        self
    }
    pub fn with_background(mut self, r: u8, g: u8, b: u8, a: u8) -> Self {
        self.colors.background = (r, g, b, a);
        self
    }

    pub fn add(&mut self, widget: Box<dyn PanelWidget>) {
        self.children.push(widget);
        self.relayout();
    }

    pub fn child(&self, index: usize) -> &dyn PanelWidget {
        &*self.children[index]
    }
    pub fn child_mut(&mut self, index: usize) -> &mut dyn PanelWidget {
        &mut *self.children[index]
    }
    pub fn child_count(&self) -> usize {
        self.children.len()
    }

    pub fn remove(&mut self, index: usize) -> Box<dyn PanelWidget> {
        let w = self.children.remove(index);
        self.relayout();
        w
    }

    fn relayout(&mut self) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        match self.orientation {
            Orientation::Horizontal => self.layout_horizontal(),
            Orientation::Vertical => self.layout_vertical(),
        }
    }

    fn layout_horizontal(&mut self) {
        let r = self.rect;
        let mut cx = r.x + self.padding;
        let mut cy = r.y + self.padding;
        let max_x = r.x + r.w - self.padding;
        let mut row_height: f32 = 0.0;

        for child in &mut self.children {
            let cr = child.rect();
            let cw = self.item_width.unwrap_or(cr.w.max(20.0));
            let ch = self.item_height.unwrap_or(cr.h.max(16.0));

            if cx + cw > max_x && cx > r.x + self.padding {
                cx = r.x + self.padding;
                cy += row_height + self.spacing;
                row_height = 0.0;
            }

            child.set_rect(LayoutRect::new(cx, cy, cw, ch));
            cx += cw + self.spacing;
            row_height = row_height.max(ch);
        }
    }

    fn layout_vertical(&mut self) {
        let r = self.rect;
        let mut cx = r.x + self.padding;
        let mut cy = r.y + self.padding;
        let max_y = r.y + r.h - self.padding;
        let mut col_width: f32 = 0.0;

        for child in &mut self.children {
            let cr = child.rect();
            let cw = self.item_width.unwrap_or(cr.w.max(20.0));
            let ch = self.item_height.unwrap_or(cr.h.max(16.0));

            if cy + ch > max_y && cy > r.y + self.padding {
                cy = r.y + self.padding;
                cx += col_width + self.spacing;
                col_width = 0.0;
            }

            child.set_rect(LayoutRect::new(cx, cy, cw, ch));
            cy += ch + self.spacing;
            col_width = col_width.max(cw);
        }
    }
}

impl PanelWidget for WrapPanel {
    /// The document tree's children — what `find_widget_mut` / `take_widget`
    /// walk, and what makes a node reachable by name however deeply nested.
    fn children_mut(&mut self) -> Vec<&mut Box<dyn PanelWidget>> {
        self.children.iter_mut().collect()
    }

    /// `removeChild` against a direct child.
    fn detach(&mut self, name: &str) -> Option<Box<dyn PanelWidget>> {
        let i = self.children.iter().position(|c| c.name() == name)?;
        Some(self.children.remove(i))
    }
    fn find_rect(&self, name: &str) -> Option<LayoutRect> {
        if self.name() == name {
            return Some(self.rect());
        }
        for child in &self.children {
            if let Some(r) = child.find_rect(name) {
                return Some(r);
            }
        }
        None
    }
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }

    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.relayout();
    }

    fn rect(&self) -> LayoutRect {
        self.rect
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        let (br, bg, bb, ba) = self.colors.background;
        if ba > 0 {
            let mut paint = Paint::default();
            paint.set_color_rgba8(br, bg, bb, ba);
            let ts = Transform::from_scale(ctx.scale, ctx.scale);
            if let Some(rect) = Rect::from_xywh(r.x, r.y, r.w, r.h) {
                ctx.pixmap.fill_rect(rect, &paint, ts, None);
            }
        }

        for child in &mut self.children {
            child.render(ctx);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        for child in self.children.iter_mut().rev() {
            if child.rect().contains(event.x, event.y) && child.handle_mouse(event) {
                return true;
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        for child in &mut self.children {
            if child.handle_key(event) {
                return true;
            }
        }
        false
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        for child in self.children.iter_mut().rev() {
            if child.rect().contains(x, y) && child.handle_scroll(delta, x, y) {
                return true;
            }
        }
        false
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        for child in self.children.iter().rev() {
            if child.rect().contains(x, y) {
                return child.cursor_at(x, y);
            }
        }
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = Vec::new();
        for child in &mut self.children {
            events.extend(child.drain_events());
        }
        events
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetEnabled(_) | WidgetCommand::SetVisible(_) => {
                for child in &mut self.children {
                    child.handle_command(cmd);
                }
                CommandValue::None
            }
            _ => CommandValue::None,
        }
    }
}
