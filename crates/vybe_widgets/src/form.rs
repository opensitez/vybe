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

use super::layout::{
    CommandValue, FocusManager, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext,
    WidgetCommand, WidgetEvent, WidgetId,
};
use std::collections::HashMap;
use tiny_skia::*;

/// A form holds a collection of controls laid out at absolute positions.
///
/// It implements `PanelWidget` so it can be nested inside containers
/// (DockPanel, SplitPanel, TabPanel, etc.).
pub struct Form {
    pub id: WidgetId,
    pub title: String,
    pub background: (u8, u8, u8, u8),
    rect: LayoutRect,
    controls: Vec<Box<dyn PanelWidget>>,
    /// Per-control `z-index`. Absent is `0`; `auto` removes the entry.
    child_z: HashMap<String, i32>,
    focus: FocusManager,
    pending_events: Vec<WidgetEvent>,
    /// Children staged for a parent whose widget isn't attached yet — the
    /// declarative (Flutter) tree assembles bottom-up (a Row nests its buttons
    /// before the Row itself is nested into a Column). Keyed by parent name.
    pending_children: HashMap<String, Vec<Box<dyn PanelWidget>>>,
    /// Index of a control that fills the form (the Flutter `runApp` root) — it
    /// is re-sized to the form's rect on every `set_rect` so its flow layout
    /// cascades once the window supplies a real size.
    fill_root: Option<usize>,
}

impl Form {
    pub fn new(title: &str) -> Self {
        Self {
            id: WidgetId::next(),
            title: title.to_string(),
            background: (240, 240, 240, 255),
            rect: LayoutRect::zero(),
            controls: Vec::new(),
            child_z: HashMap::new(),
            focus: FocusManager::new(),
            pending_events: Vec::new(),
            pending_children: HashMap::new(),
            fill_root: None,
        }
    }

    /// Stage a control into the widget tree (declarative/Flutter path). The
    /// host creates `widget` (via `make_widget`) and tells us its `name`, its
    /// `parent` name, and whether the parent is the form itself. We first
    /// adopt any children that were waiting on `name` (bottom-up assembly),
    /// then either fill+attach to the form (parent is the form) or park the
    /// finished subtree until `parent`'s own widget is staged. vybe_widgets
    /// owns the tree + layout; the host stays a thin bridge.
    pub fn stage_control(
        &mut self,
        name: &str,
        mut widget: Box<dyn PanelWidget>,
        parent: &str,
        parent_is_form: bool,
    ) {
        let _ = name;
        if parent_is_form {
            // Fill the form so the root's relayout cascades down the subtree.
            // Recorded as `fill_root` so a later `set_rect` (real window size)
            // re-fills and re-lays-out.
            widget.set_rect(LayoutRect::new(
                self.rect.x,
                self.rect.y,
                self.rect.w,
                self.rect.h,
            ));
            self.fill_root = Some(self.controls.len());
            self.controls.push(widget);
        } else {
            // Top-down: the parent control already exists in the tree — add the
            // child into it by name.
            let mut child = Some(widget);
            for ctrl in &mut self.controls {
                if let Some(w) = child.take() {
                    child = ctrl.add_child_to(parent, w);
                }
                if child.is_none() {
                    break;
                }
            }
            // If `child` is still Some, the parent wasn't found — drop it.
        }
    }

    /// Drop every control + staged subtree — used by the Flutter `setState`
    /// rebuild, which re-realizes the widget tree from scratch (state persists
    /// separately, in the Dart runtime's per-type State cache).
    pub fn clear_controls(&mut self) {
        self.controls.clear();
        self.pending_children.clear();
        self.fill_root = None;
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

    /// Add an already-boxed control at an absolute position.
    pub fn add_boxed_control(
        &mut self,
        mut widget: Box<dyn PanelWidget>,
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
        self.controls.push(widget);
    }

    /// Number of controls on this form.
    pub fn control_count(&self) -> usize {
        self.controls.len()
    }

    /// Dump all widgets' names, positions and types for debugging.
    pub fn debug_dump(&self) {
        eprintln!(
            "[WIDGET-DUMP] Form '{}' bg={:?} controls={}:",
            self.title,
            self.background,
            self.controls.len()
        );
        for (i, w) in self.controls.iter().enumerate() {
            let r = w.rect();
            eprintln!(
                "  [{}] name='{}' rect=({:.0},{:.0} {:.0}x{:.0})",
                i,
                w.name(),
                r.x,
                r.y,
                r.w,
                r.h
            );
        }
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

    /// Get the layout rect of a named control (recursively searching containers).
    pub fn get_control_rect(&self, name: &str) -> Option<LayoutRect> {
        for w in &self.controls {
            if let Some(r) = w.find_rect(name) {
                return Some(r);
            }
        }
        None
    }

    /// Render host overlays on top of controls (owner-draw).
    /// Put `controls` into paint order.
    ///
    /// Shares `layout::paint_order` with the flow container rather than sorting
    /// here: there is one stacking rule and it should have one implementation.
    /// Every body-level control is positioned, so the negative-z bucket is what
    /// puts a `z-index: -1` control behind its siblings.
    fn reorder_for_paint(&mut self) {
        let order = crate::layout::paint_order(
            self.controls.len(),
            |_| true,
            |i| {
                self.child_z
                    .get(self.controls[i].name())
                    .copied()
                    .unwrap_or(0)
            },
        );
        let mut taken: Vec<Option<Box<dyn PanelWidget>>> =
            self.controls.drain(..).map(Some).collect();
        self.controls = order.into_iter().filter_map(|i| taken[i].take()).collect();
    }

    // `render_overlays` was here and is deleted.
    //
    // It replayed a name-keyed map of recordings onto the form, and its ONLY
    // caller was `gui_capture.rs` — never the window runner. So a drawing that
    // reached the overlay map appeared in `--capture` and nowhere else, which
    // made the screenshot the least trustworthy thing in the pipeline.
    //
    // A `<canvas>` is an element in the document now. It has a rect, it paints
    // in tree order with every other control, and a capture and a window show
    // the same pixels because they run the same code — not because two paths
    // were kept in agreement.

    /// Mutable access to the form's child controls slice. Used by the
    /// host bridge to walk widgets and downcast them when looking up a
    /// `Canvas` widget by name.
    pub fn controls_mut(&mut self) -> &mut [Box<dyn PanelWidget>] {
        &mut self.controls
    }

    /// Read-only access to the form's child controls slice.
    pub fn controls(&self) -> &[Box<dyn PanelWidget>] {
        &self.controls
    }

    /// Move focus to the next focusable control (Tab key).
    pub fn focus_next(&mut self) {
        self.focus.focus_next(&mut self.controls);
    }

    /// Move focus to the previous focusable control (Shift+Tab).
    pub fn focus_prev(&mut self) {
        self.focus.focus_prev(&mut self.controls);
    }

    /// Recalculate control positions when the form's rect changes.
    #[allow(dead_code)]
    fn relayout_controls(&mut self) {
        // Controls keep their relative positions within the form.
        // When the form moves, we'd need stored offsets. For now,
        // controls are positioned absolutely via add_control.
    }

    /// Send a command to a child control by name.
    pub fn send_command(&mut self, name: &str, cmd: &WidgetCommand) -> CommandValue {
        self.focus.send_command(&mut self.controls, name, cmd)
    }

    /// Send a command to a child control by WidgetId.
    pub fn send_command_by_id(&mut self, id: WidgetId, cmd: &WidgetCommand) -> CommandValue {
        self.focus.send_command_by_id(&mut self.controls, id, cmd)
    }

    /// Broadcast a command to all child controls.
    pub fn broadcast_command(&mut self, cmd: &WidgetCommand) -> Vec<(WidgetId, CommandValue)> {
        self.focus.broadcast_command(&mut self.controls, cmd)
    }
}

impl PanelWidget for Form {
    /// The document tree's children — what `find_widget_mut` / `take_widget`
    /// walk, and what makes a node reachable by name however deeply nested.
    fn children_mut(&mut self) -> Vec<&mut Box<dyn PanelWidget>> {
        self.controls.iter_mut().collect()
    }

    /// `removeChild` against a direct child.
    fn detach(&mut self, name: &str) -> Option<Box<dyn PanelWidget>> {
        let i = self.controls.iter().position(|c| c.name() == name)?;
        Some(self.controls.remove(i))
    }
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
        // A Flutter `runApp` root fills the form and re-lays-out its subtree
        // whenever the form is (re)sized.
        if let Some(i) = self.fill_root {
            if let Some(root) = self.controls.get_mut(i) {
                root.set_rect(rect);
            }
        }
    }

    fn rect(&self) -> LayoutRect {
        self.rect
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        // Fill background
        let ts = Transform::from_scale(ctx.scale, ctx.scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let (br, bg, bb, ba) = self.background;
        paint.set_color_rgba8(br, bg, bb, ba);
        if let Some(rect) = Rect::from_xywh(r.x, r.y, r.w, r.h) {
            ctx.pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Render all controls (with focus ring on focused one)
        self.focus.render_all(&mut self.controls, ctx);
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            return false;
        }
        self.focus.handle_mouse(&mut self.controls, event);
        true // Consume to prevent fall-through
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        self.focus.handle_key(&mut self.controls, event)
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        if !self.rect.contains(x, y) {
            return false;
        }
        for ctrl in self.controls.iter_mut().rev() {
            if ctrl.handle_scroll(delta, x, y) {
                return true;
            }
        }
        false
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        if !self.rect.contains(x, y) {
            return winit::window::CursorIcon::Default;
        }
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

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetText(t) => {
                self.title = t.clone();
                CommandValue::None
            }
            WidgetCommand::GetText => CommandValue::Text(self.title.clone()),
            // `z-index` on a body-level control. The form has no flow layout —
            // every control it holds carries its own coordinates, so all of
            // them are positioned and the paint order is entirely the sort.
            //
            // Applied by REORDERING `controls` rather than by ordering the
            // render loop, so hit-testing follows painting: the box drawn on
            // top is the box that gets the click, which is the one thing a
            // separate paint order would silently get wrong.
            WidgetCommand::Custom(name, value) if name == "SetChildZ" => {
                if let CommandValue::Text(spec) = value {
                    if let Some((child, z)) = spec.rsplit_once('=') {
                        match z.trim().parse::<i32>() {
                            Ok(z) => {
                                self.child_z.insert(child.to_string(), z);
                            }
                            Err(_) => {
                                self.child_z.remove(child);
                            }
                        }
                        self.reorder_for_paint();
                    }
                }
                CommandValue::None
            }
            // `SetChildMargin` deliberately has no arm.
            //
            // A margin is space a container puts BETWEEN siblings, and the form
            // has no flow: every control it holds carries its own coordinates.
            // The other half of what a margin does — displacing a positioned
            // box from its own offsets — belongs to the offsets, and the
            // document applies it when it writes `left`/`top`. So there is
            // nothing left here to act on, and inventing a flow for the body to
            // have something to do with it would be the wrong answer twice.
            _ => CommandValue::None,
        }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = std::mem::take(&mut self.pending_events);
        events.extend(self.focus.drain_all_events(&mut self.controls));
        events
    }
}
