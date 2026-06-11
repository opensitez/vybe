//! FlowLayoutPanel — arranges children in a flow (left-to-right, wrapping to next row).
//!
//! Like WinForms FlowLayoutPanel: children are placed sequentially; when a child
//! would overflow the current row it wraps to the next row.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetCommand,
    WidgetEvent, WidgetId,
};
use tiny_skia::*;

/// Flow direction for child arrangement.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FlowDirection {
    LeftToRight,
    TopDown,
}

pub struct FlowLayoutPanel {
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    pub flow_direction: FlowDirection,
    /// Spacing between children in pixels.
    pub spacing: f32,
    /// Padding inside the panel edges.
    pub padding: f32,
    /// Whether children wrap to the next row/column when they exceed the panel size.
    pub wrap_contents: bool,
    rect: LayoutRect,
    children: Vec<Box<dyn PanelWidget>>,
}

impl FlowLayoutPanel {
    pub fn new() -> Self {
        Self {
            width: 300.0,
            height: 200.0,
            colors: WidgetColors {
                background: (250, 250, 250, 255),
                border: (180, 180, 180, 255),
                ..WidgetColors::default()
            },
            id: WidgetId::next(),
            name: String::new(),
            flow_direction: FlowDirection::LeftToRight,
            spacing: 4.0,
            padding: 4.0,
            wrap_contents: true,
            rect: LayoutRect::zero(),
            children: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }
    pub fn with_direction(mut self, dir: FlowDirection) -> Self {
        self.flow_direction = dir;
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
    pub fn with_wrap(mut self, wrap: bool) -> Self {
        self.wrap_contents = wrap;
        self
    }

    /// Add a child widget. Triggers relayout.
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

    /// Remove a child by index. Triggers relayout.
    pub fn remove(&mut self, index: usize) -> Box<dyn PanelWidget> {
        let w = self.children.remove(index);
        self.relayout();
        w
    }

    /// Arrange children according to flow direction.
    fn relayout(&mut self) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        match self.flow_direction {
            FlowDirection::LeftToRight => self.layout_left_to_right(),
            FlowDirection::TopDown => self.layout_top_down(),
        }
    }

    fn layout_left_to_right(&mut self) {
        let r = self.rect;
        let mut cx = r.x + self.padding;
        let mut cy = r.y + self.padding;
        let mut row_height: f32 = 0.0;
        let max_x = r.x + r.w - self.padding;

        for child in &mut self.children {
            let cr = child.rect();
            let cw = cr.w.max(20.0); // minimum child width
            let ch = cr.h.max(16.0); // minimum child height

            // Wrap to next row if needed
            if self.wrap_contents && cx + cw > max_x && cx > r.x + self.padding {
                cx = r.x + self.padding;
                cy += row_height + self.spacing;
                row_height = 0.0;
            }

            child.set_rect(LayoutRect::new(cx, cy, cw, ch));
            cx += cw + self.spacing;
            row_height = row_height.max(ch);
        }
    }

    fn layout_top_down(&mut self) {
        let r = self.rect;
        let mut cx = r.x + self.padding;
        let mut cy = r.y + self.padding;
        let mut col_width: f32 = 0.0;
        let max_y = r.y + r.h - self.padding;

        for child in &mut self.children {
            let cr = child.rect();
            let cw = cr.w.max(20.0);
            let ch = cr.h.max(16.0);

            // Wrap to next column if needed
            if self.wrap_contents && cy + ch > max_y && cy > r.y + self.padding {
                cy = r.y + self.padding;
                cx += col_width + self.spacing;
                col_width = 0.0;
            }

            child.set_rect(LayoutRect::new(cx, cy, cw, ch));
            cy += ch + self.spacing;
            col_width = col_width.max(cw);
        }
    }

    /// Paint — light background with dashed border.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Light background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Dashed border
        let (r, g, b, a) = self.colors.border;
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        stroke.dash = StrokeDash::new(vec![4.0, 3.0], 0.0);

        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

impl PanelWidget for FlowLayoutPanel {
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
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
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
