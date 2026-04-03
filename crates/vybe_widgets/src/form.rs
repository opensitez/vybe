//! Form — a window-like container that holds controls and manages focus.
//!
//! ```ignore
//! use vybe_widgets::*;
//!
//! let mut form = Form::new("My App");
//! form.add_control(Button::new("OK").with_name("btn_ok"), 100.0, 200.0, 80.0, 28.0);
//! form.add_control(TextInput::new().with_name("name").with_placeholder("Name"), 10.0, 10.0, 200.0, 24.0);
//! form.add_control(Label::new("Hello"), 10.0, 40.0, 100.0, 20.0);
//! form.add_control(Checkbox::new("Accept").with_name("accept"), 10.0, 70.0, 150.0, 20.0);
//!
//! // In your Application impl, just call:
//! //   form.render(&mut ctx);
//! //   form.handle_mouse(&event);
//! //   form.handle_key(&event);
//! //   for ev in form.drain_events() { match ev { ... } }
//! ```

use tiny_skia::*;
use super::layout::{
    LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton,
    KeyEvent, RenderContext, PanelWidget, WidgetEvent,
};

/// A form holds a collection of controls laid out at absolute positions.
///
/// It implements `PanelWidget` so it can be nested inside containers
/// (DockPanel, SplitPanel, TabPanel, etc.).
pub struct Form {
    pub title: String,
    pub background: (u8, u8, u8, u8),
    rect: LayoutRect,
    controls: Vec<Box<dyn PanelWidget>>,
    focused_index: Option<usize>,
    pending_events: Vec<WidgetEvent>,
}

impl Form {
    pub fn new(title: &str) -> Self {
        Self {
            title: title.to_string(),
            background: (240, 240, 240, 255),
            rect: LayoutRect::zero(),
            controls: Vec::new(),
            focused_index: None,
            pending_events: Vec::new(),
        }
    }

    /// Add a control at an absolute position within the form.
    pub fn add_control<W: PanelWidget + 'static>(
        &mut self,
        mut widget: W,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
    ) {
        widget.set_rect(LayoutRect::new(
            self.rect.x + x,
            self.rect.y + y,
            width,
            height,
        ));
        self.controls.push(Box::new(widget));
    }

    /// Number of controls on this form.
    pub fn control_count(&self) -> usize {
        self.controls.len()
    }

    /// Get a reference to a control by index.
    pub fn control(&self, index: usize) -> Option<&dyn PanelWidget> {
        self.controls.get(index).map(|c| c.as_ref())
    }

    /// Get a mutable reference to a control by index.
    pub fn control_mut(&mut self, index: usize) -> Option<&mut (dyn PanelWidget + 'static)> {
        match self.controls.get_mut(index) {
            Some(c) => Some(c.as_mut()),
            None => None,
        }
    }

    /// Move focus to the next focusable control (Tab key).
    pub fn focus_next(&mut self) {
        let count = self.controls.len();
        if count == 0 { return; }
        let start = self.focused_index.map(|i| i + 1).unwrap_or(0);
        for offset in 0..count {
            let idx = (start + offset) % count;
            if self.controls[idx].focusable() {
                self.focused_index = Some(idx);
                return;
            }
        }
    }

    /// Move focus to the previous focusable control (Shift+Tab).
    pub fn focus_prev(&mut self) {
        let count = self.controls.len();
        if count == 0 { return; }
        let start = self.focused_index.unwrap_or(0);
        for offset in 1..=count {
            let idx = (start + count - offset) % count;
            if self.controls[idx].focusable() {
                self.focused_index = Some(idx);
                return;
            }
        }
    }

    /// Recalculate control positions when the form's rect changes.
    fn relayout_controls(&mut self) {
        // Controls keep their relative positions within the form.
        // When the form moves, we'd need stored offsets. For now,
        // controls are positioned absolutely via add_control.
    }
}

impl PanelWidget for Form {
    fn set_rect(&mut self, rect: LayoutRect) {
        let dx = rect.x - self.rect.x;
        let dy = rect.y - self.rect.y;
        self.rect = rect;
        // Shift all controls by the delta
        if dx != 0.0 || dy != 0.0 {
            for ctrl in &mut self.controls {
                let cr = ctrl.rect();
                ctrl.set_rect(LayoutRect::new(cr.x + dx, cr.y + dy, cr.w, cr.h));
            }
        }
    }

    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }

        // Fill background
        let ts = Transform::from_scale(ctx.scale, ctx.scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let (br, bg, bb, ba) = self.background;
        paint.set_color_rgba8(br, bg, bb, ba);
        if let Some(rect) = Rect::from_xywh(r.x, r.y, r.w, r.h) {
            ctx.pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Render all controls
        for ctrl in &mut self.controls {
            ctrl.render(ctx);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) { return false; }

        // On click, update focus to the clicked control
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            let mut new_focus = None;
            for (i, ctrl) in self.controls.iter().enumerate() {
                if ctrl.rect().contains(event.x, event.y) && ctrl.focusable() {
                    new_focus = Some(i);
                    break;
                }
            }
            self.focused_index = new_focus;
        }

        // Route to controls (topmost / last added wins)
        for ctrl in self.controls.iter_mut().rev() {
            if ctrl.handle_mouse(event) {
                return true;
            }
        }
        true // Consume to prevent fall-through
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        use winit::keyboard::{Key, NamedKey};
        use winit::event::ElementState;

        // Tab / Shift+Tab for focus navigation
        if event.state == ElementState::Pressed {
            if let Key::Named(NamedKey::Tab) = &event.key_without_modifiers {
                if event.shift { self.focus_prev(); } else { self.focus_next(); }
                return true;
            }
        }

        // Route to focused control
        if let Some(idx) = self.focused_index {
            if let Some(ctrl) = self.controls.get_mut(idx) {
                if ctrl.handle_key(event) {
                    return true;
                }
            }
        }
        false
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        if !self.rect.contains(x, y) { return false; }
        for ctrl in self.controls.iter_mut().rev() {
            if ctrl.handle_scroll(delta, x, y) {
                return true;
            }
        }
        false
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        if !self.rect.contains(x, y) { return winit::window::CursorIcon::Default; }
        for ctrl in self.controls.iter().rev() {
            let r = ctrl.rect();
            if r.contains(x, y) {
                let c = ctrl.cursor_at(x, y);
                if c != winit::window::CursorIcon::Default {
                    return c;
                }
            }
        }
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = std::mem::take(&mut self.pending_events);
        for ctrl in &mut self.controls {
            events.extend(ctrl.drain_events());
        }
        events
    }
}
