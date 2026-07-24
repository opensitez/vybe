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
    /// This panel's flex weight when it is itself a child of another flex
    /// container (0 = fixed/natural size). Default 1.
    pub flex: f32,
    /// Fixed main-axis size (height in TopDown, width in LeftToRight) used when
    /// `flex == 0`. Default ~a toolbar height.
    pub fixed_main: f32,
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
            flow_direction: FlowDirection::TopDown,
            spacing: 4.0,
            padding: 4.0,
            wrap_contents: true,
            flex: 1.0,
            fixed_main: 44.0,
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

    // Flutter Row: children side by side. Fixed-flex children keep their
    // `fixed_main` width; flex children share the leftover by weight. Full
    // height.
    fn layout_left_to_right(&mut self) {
        let r = self.rect;
        let n = self.children.len();
        if n == 0 {
            return;
        }
        let inner_w = (r.w - 2.0 * self.padding).max(0.0);
        let inner_h = (r.h - 2.0 * self.padding).max(0.0);
        let gaps = self.spacing * (n as f32 - 1.0);
        let (total_flex, fixed) = self.flex_totals();
        let leftover = (inner_w - gaps - fixed).max(0.0);
        let mut cx = r.x + self.padding;
        let cy = r.y + self.padding;
        for child in &mut self.children {
            let f = child.layout_flex();
            let cw = if f <= 0.0 {
                Self::child_fixed(child.as_ref())
            } else if total_flex > 0.0 {
                leftover * f / total_flex
            } else {
                0.0
            };
            child.set_rect(LayoutRect::new(cx, cy, cw, inner_h));
            cx += cw + self.spacing;
        }
    }

    // Flutter Column: children stacked top to bottom. Fixed-flex children keep
    // their `fixed_main` height; flex children share the leftover by weight.
    // Full width.
    fn layout_top_down(&mut self) {
        let r = self.rect;
        let n = self.children.len();
        if n == 0 {
            return;
        }
        let inner_w = (r.w - 2.0 * self.padding).max(0.0);
        let inner_h = (r.h - 2.0 * self.padding).max(0.0);
        let gaps = self.spacing * (n as f32 - 1.0);
        let (total_flex, fixed) = self.flex_totals();
        let leftover = (inner_h - gaps - fixed).max(0.0);
        let cx = r.x + self.padding;
        let mut cy = r.y + self.padding;
        for child in &mut self.children {
            let f = child.layout_flex();
            let ch = if f <= 0.0 {
                Self::child_fixed(child.as_ref())
            } else if total_flex > 0.0 {
                leftover * f / total_flex
            } else {
                0.0
            };
            child.set_rect(LayoutRect::new(cx, cy, inner_w, ch));
            cy += ch + self.spacing;
        }
    }

    /// (sum of flex weights, sum of fixed children's main-axis sizes).
    fn flex_totals(&self) -> (f32, f32) {
        let mut total_flex = 0.0;
        let mut fixed = 0.0;
        for child in &self.children {
            let f = child.layout_flex();
            if f <= 0.0 {
                fixed += Self::child_fixed(child.as_ref());
            } else {
                total_flex += f;
            }
        }
        (total_flex, fixed)
    }

    /// The fixed main-axis size of a flex-0 child (a toolbar-height bar).
    fn child_fixed(_child: &dyn PanelWidget) -> f32 {
        44.0
    }

    /// Paint — a Flutter layout container is transparent (no chrome); only its
    /// children paint. Kept as a no-op so Column/Row/Scaffold don't draw the
    /// WinForms-style dashed panel border.
    pub fn paint(&self, _pixmap: &mut Pixmap, _x: f32, _y: f32, _scale: f32) {}

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

    fn add_child(&mut self, child: Box<dyn PanelWidget>) -> Option<Box<dyn PanelWidget>> {
        self.add(child); // pushes + relayout()
        None
    }

    fn send_command_named(
        &mut self,
        name: &str,
        cmd: &WidgetCommand,
    ) -> Option<CommandValue> {
        if self.name == name {
            return Some(self.handle_command(cmd));
        }
        for child in &mut self.children {
            if let Some(result) = child.send_command_named(name, cmd) {
                return Some(result);
            }
        }
        None
    }

    fn add_child_to(
        &mut self,
        parent_name: &str,
        child: Box<dyn PanelWidget>,
    ) -> Option<Box<dyn PanelWidget>> {
        if self.name == parent_name {
            return self.add_child(child);
        }
        let mut child = Some(child);
        for existing in &mut self.children {
            if let Some(c) = child.take() {
                child = existing.add_child_to(parent_name, c);
            }
            if child.is_none() {
                break;
            }
        }
        child
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

    fn layout_flex(&self) -> f32 {
        self.flex
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetFlex(f) => {
                self.flex = *f;
                CommandValue::None
            }
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
