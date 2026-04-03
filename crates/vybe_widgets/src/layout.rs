//! Layout primitives and the `PanelWidget` trait.
//!
//! This module provides the foundation for the vybe GUI toolkit:
//! - `LayoutRect` — rectangle with hit-testing and subdivision helpers
//! - `MouseEvent` / `KeyEvent` — unified input events
//! - `RenderContext` — bundle of rendering resources
//! - `PanelWidget` trait — implemented by all toolkit panels and containers

use cosmic_text::{FontSystem, SwashCache, Attrs, Buffer, Color, Family, Metrics, Shaping};
use tiny_skia::{Pixmap, PixmapPaint, Transform, ColorU8};
use winit::window::CursorIcon;
use std::sync::atomic::{AtomicU64, Ordering};

// ── WidgetId ───────────────────────────────────────────────────────────

/// A unique, lightweight identifier for a widget instance.
///
/// Created via `WidgetId::next()` which returns a globally unique id.
/// Zero is reserved as the "null" / uninitialized sentinel.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct WidgetId(pub u64);

static NEXT_WIDGET_ID: AtomicU64 = AtomicU64::new(1);

impl WidgetId {
    /// The null / uninitialized id.
    pub const NONE: WidgetId = WidgetId(0);

    /// Allocate the next globally-unique widget id.
    pub fn next() -> Self {
        WidgetId(NEXT_WIDGET_ID.fetch_add(1, Ordering::Relaxed))
    }

    /// Whether this is the null id.
    pub fn is_none(self) -> bool { self.0 == 0 }
}

impl Default for WidgetId {
    fn default() -> Self { Self::NONE }
}

// ── LayoutRect ─────────────────────────────────────────────────────────

/// A rectangle in logical (unscaled) coordinates used for layout.
#[derive(Clone, Copy, Debug, Default)]
pub struct LayoutRect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

impl LayoutRect {
    pub fn new(x: f32, y: f32, w: f32, h: f32) -> Self {
        Self { x, y, w, h }
    }

    pub fn zero() -> Self {
        Self { x: 0.0, y: 0.0, w: 0.0, h: 0.0 }
    }

    /// Test whether a point lies inside this rectangle.
    pub fn contains(&self, px: f32, py: f32) -> bool {
        px >= self.x && px < self.x + self.w && py >= self.y && py < self.y + self.h
    }

    /// Right edge x coordinate.
    pub fn right(&self) -> f32 { self.x + self.w }

    /// Bottom edge y coordinate.
    pub fn bottom(&self) -> f32 { self.y + self.h }

    /// Take a sub-rect from the left edge, shrinking `self`.
    pub fn take_left(&mut self, width: f32) -> LayoutRect {
        let w = width.min(self.w).max(0.0);
        let r = LayoutRect::new(self.x, self.y, w, self.h);
        self.x += w;
        self.w = (self.w - w).max(0.0);
        r
    }

    /// Take a sub-rect from the right edge, shrinking `self`.
    pub fn take_right(&mut self, width: f32) -> LayoutRect {
        let w = width.min(self.w).max(0.0);
        self.w = (self.w - w).max(0.0);
        LayoutRect::new(self.x + self.w, self.y, w, self.h)
    }

    /// Take a sub-rect from the top edge, shrinking `self`.
    pub fn take_top(&mut self, height: f32) -> LayoutRect {
        let h = height.min(self.h).max(0.0);
        let r = LayoutRect::new(self.x, self.y, self.w, h);
        self.y += h;
        self.h = (self.h - h).max(0.0);
        r
    }

    /// Take a sub-rect from the bottom edge, shrinking `self`.
    pub fn take_bottom(&mut self, height: f32) -> LayoutRect {
        let h = height.min(self.h).max(0.0);
        self.h = (self.h - h).max(0.0);
        LayoutRect::new(self.x, self.y + self.h, self.w, h)
    }
}

// ── Input Events ───────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum MouseButton {
    Left,
    Right,
    Middle,
}

#[derive(Clone, Copy, Debug)]
pub enum MouseEventKind {
    Press(MouseButton),
    Release(MouseButton),
    Move,
    Scroll(f32),
}

/// A mouse event with absolute logical coordinates.
#[derive(Clone, Copy, Debug)]
pub struct MouseEvent {
    pub x: f32,
    pub y: f32,
    pub kind: MouseEventKind,
    pub cmd: bool,
    pub shift: bool,
    pub alt: bool,
}

/// A keyboard event with modifier state.
#[derive(Clone, Debug)]
pub struct KeyEvent {
    pub logical_key: winit::keyboard::Key,
    pub key_without_modifiers: winit::keyboard::Key,
    pub state: winit::event::ElementState,
    pub cmd: bool,
    pub shift: bool,
    pub alt: bool,
    pub text: Option<String>,
}

// ── Render Context ─────────────────────────────────────────────────────

/// Bundle of resources needed for rendering.
pub struct RenderContext<'a> {
    pub pixmap: &'a mut Pixmap,
    pub font_system: &'a mut FontSystem,
    pub swash_cache: &'a mut SwashCache,
    pub scale: f32,
}

impl<'a> RenderContext<'a> {
    /// Draw monospace UI text at physical pixel coordinates.
    pub fn draw_text(&mut self, text: &str, x: f32, y: f32, r: u8, g: u8, b: u8, a: u8) {
        let col = Color::rgba(r, g, b, a);
        let mut lab = Buffer::new(self.font_system, Metrics::new(14.0, 20.0).scale(self.scale));
        lab.set_text(self.font_system, text, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None);
        lab.shape_until_scroll(self.font_system, false);
        for run in lab.layout_runs() {
            for g in run.glyphs {
                let pg = g.physical((x, y + run.line_y), 1.0);
                if let Some(img) = self.swash_cache.get_image(self.font_system, pg.cache_key) {
                    if let Some(mut p) = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)) {
                        let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a());
                        for (idx, &al) in img.data.iter().enumerate() {
                            let af = (al as f32 / 255.0) * (ca as f32 / 255.0);
                            p.pixels_mut()[idx] = ColorU8::from_rgba(
                                (cr as f32 * af) as u8, (cg as f32 * af) as u8,
                                (cb as f32 * af) as u8, (255.0 * af) as u8,
                            ).premultiply();
                        }
                        self.pixmap.draw_pixmap(pg.x + img.placement.left, pg.y - img.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                    }
                }
            }
        }
    }

    /// Draw monospace UI text with a custom font size at physical pixel coordinates.
    pub fn draw_text_sized(&mut self, text: &str, x: f32, y: f32, font_size: f32, r: u8, g: u8, b: u8, a: u8) {
        let col = Color::rgba(r, g, b, a);
        let mut lab = Buffer::new(self.font_system, Metrics::new(font_size, font_size * 1.5).scale(self.scale));
        lab.set_text(self.font_system, text, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None);
        lab.shape_until_scroll(self.font_system, false);
        for run in lab.layout_runs() {
            for g in run.glyphs {
                let pg = g.physical((x, y + run.line_y), 1.0);
                if let Some(img) = self.swash_cache.get_image(self.font_system, pg.cache_key) {
                    if let Some(mut p) = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)) {
                        let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a());
                        for (idx, &al) in img.data.iter().enumerate() {
                            let af = (al as f32 / 255.0) * (ca as f32 / 255.0);
                            p.pixels_mut()[idx] = ColorU8::from_rgba(
                                (cr as f32 * af) as u8, (cg as f32 * af) as u8,
                                (cb as f32 * af) as u8, (255.0 * af) as u8,
                            ).premultiply();
                        }
                        self.pixmap.draw_pixmap(pg.x + img.placement.left, pg.y - img.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                    }
                }
            }
        }
    }
}

// ── Cursor Motion ──────────────────────────────────────────────────────

/// Direction for cursor movement — abstracts cosmic_text::Motion.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum CursorMotion {
    Left,
    Right,
    Up,
    Down,
    Home,
    End,
    BufferStart,
    BufferEnd,
    LeftWord,
    RightWord,
}

// ── Dock ───────────────────────────────────────────────────────────────

/// Dock position for children of a `DockPanel`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Dock {
    Left,
    Right,
    Top,
    Bottom,
    Fill,
}

// ── Widget Events ──────────────────────────────────────────────────────

/// A dynamic value that can be exchanged via commands.
#[derive(Clone, Debug)]
pub enum CommandValue {
    /// No value / not applicable.
    None,
    /// A string value (text, label, path, etc.).
    Text(String),
    /// A numeric value (slider position, counter value, etc.).
    Number(f64),
    /// A boolean value (checked, enabled, visible, etc.).
    Bool(bool),
    /// An index value (selected item, tab index, etc.).
    Index(usize),
    /// An RGBA colour.
    Color(u8, u8, u8, u8),
}

/// Tri-state check state for checkboxes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CheckState {
    Unchecked,
    Checked,
    Indeterminate,
}

impl CheckState {
    /// Cycle to the next state on toggle: Unchecked -> Checked -> Unchecked.
    /// Indeterminate is only set programmatically and cycles to Checked on toggle.
    pub fn toggle(self) -> Self {
        match self {
            CheckState::Unchecked => CheckState::Checked,
            CheckState::Checked => CheckState::Unchecked,
            CheckState::Indeterminate => CheckState::Checked,
        }
    }

    pub fn is_checked(self) -> bool {
        self == CheckState::Checked
    }
}

/// Selection mode for list-based controls.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SelectionMode {
    /// Only one item can be selected at a time (default).
    Single,
    /// Multiple items can be toggled independently with simple clicks.
    MultiSimple,
    /// Extended multi-select: Ctrl+click toggles, Shift+click selects range.
    MultiExtended,
}

/// Text alignment for labels and similar widgets.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextAlign {
    Left,
    Center,
    Right,
}

/// Commands sent from the host application **to** a widget.
#[derive(Clone, Debug)]
pub enum WidgetCommand {
    /// Set the widget's text / label content.
    SetText(String),
    /// Get the widget's current text.
    GetText,
    /// Set a numeric value (slider position, progress, numeric up-down, etc.).
    SetValue(f64),
    /// Get the widget's current value.
    GetValue,
    /// Enable or disable the widget.
    SetEnabled(bool),
    /// Show or hide the widget.
    SetVisible(bool),
    /// Request focus for this widget.
    Focus,
    /// Set the selected index (list, combo, tabs, etc.).
    SetSelectedIndex(usize),
    /// Set checked / selected state (checkbox, radio).
    SetChecked(bool),
    /// Append an item to the widget's list.
    AddItem(String),
    /// Remove item at index.
    RemoveItem(usize),
    /// Remove all items.
    ClearItems,
    /// Custom command with a string key and arbitrary payload.
    Custom(String, CommandValue),
}

/// Events emitted by widgets back to the host application.
///
/// Containers collect these; the host matches on them to handle callbacks
/// without knowing widget layout details.
#[derive(Clone, Debug)]
pub enum WidgetEvent {
    /// A tab was selected. Payload: tab index.
    TabChanged(usize),
    /// A tab close button was clicked. Payload: tab index.
    TabCloseRequested(usize),
    /// A status-bar section was clicked. Payload: section click_id.
    StatusBarClick(String),
    /// A tree-view item was opened. Payload: file path.
    TreeItemOpened(String),
    /// A dropdown item was selected. Payload: dropdown id, item index.
    DropdownSelected(String, usize),
    /// A menu action was selected. Payload: action id string.
    MenuAction(String),
    /// A button was clicked. Payload: button name.
    ButtonClicked(String),
    /// A checkbox was toggled. Payload: checkbox name, new checked state.
    CheckboxToggled(String, bool),
    /// A text field value changed. Payload: field name, new text.
    TextChanged(String, String),
    /// A color was picked. Payload: widget name, hex string (e.g. "#FF0000").
    ColorChanged(String, String),
    /// A radio button was selected. Payload: name, selected state.
    RadioSelected(String, bool),
    /// A slider value changed. Payload: name, new value (0.0..1.0).
    SliderChanged(String, f32),
    /// A list box item was selected. Payload: name, selected index.
    ListBoxSelected(String, usize),
    /// A list view item was selected. Payload: name, selected index.
    ListViewSelected(String, usize),
    /// A numeric up-down value changed. Payload: name, new value.
    NumericChanged(String, f64),
    /// A scroll bar position changed. Payload: name, position (0.0..1.0).
    ScrollChanged(String, f32),
    /// A calendar date was selected. Payload: name, day of month.
    CalendarDateSelected(String, u32),
    /// A menu strip item was clicked. Payload: name, item index.
    MenuItemClicked(String, usize),
    /// A context menu item was clicked. Payload: name, item index.
    ContextMenuItemClicked(String, usize),
    /// A tool strip button was clicked. Payload: name, item index.
    ToolStripItemClicked(String, usize),
    /// A split container divider was moved. Payload: name, new position (0.0..1.0).
    SplitMoved(String, f32),
    /// A link label was clicked. Payload: name.
    LinkClicked(String),
    /// A select/combobox item was chosen. Payload: name, selected index.
    SelectChanged(String, usize),
    /// A tab control tab was selected. Payload: name, tab index.
    TabControlChanged(String, usize),
    /// Generic named action (for custom widgets).
    Action(String),
    /// Mouse entered a widget. Payload: widget name.
    MouseEnter(String),
    /// Mouse left a widget. Payload: widget name.
    MouseLeave(String),
}

// ── PanelWidget Trait ──────────────────────────────────────────────────

/// Core trait for widgets in the toolkit layout system.
///
/// All panel containers and leaf widgets implement this trait.
/// Coordinates are in logical (unscaled) pixels.
pub trait PanelWidget {
    /// Assign a layout rectangle to this widget.
    fn set_rect(&mut self, rect: LayoutRect);

    /// Get the current layout rectangle.
    fn rect(&self) -> LayoutRect;

    /// Render the widget into the pixmap.
    fn render(&mut self, ctx: &mut RenderContext);

    /// Handle a mouse event. Returns `true` if the event was consumed.
    fn handle_mouse(&mut self, event: &MouseEvent) -> bool;

    /// Handle a key event. Returns `true` if the event was consumed.
    fn handle_key(&mut self, event: &KeyEvent) -> bool;

    /// Handle a scroll event at (x, y) with the given delta.
    /// Returns `true` if the event was consumed.
    fn handle_scroll(&mut self, _delta: f32, _x: f32, _y: f32) -> bool { false }

    /// Return the desired cursor icon for the given position.
    fn cursor_at(&self, _x: f32, _y: f32) -> CursorIcon { CursorIcon::Default }

    /// Drain any pending events from this widget (and its children).
    fn drain_events(&mut self) -> Vec<WidgetEvent> { Vec::new() }

    /// Whether this widget wants keyboard focus.
    fn focusable(&self) -> bool { false }

    /// Widget name (used by FocusManager for identification).
    fn name(&self) -> &str { "" }

    /// Notify the widget that it gained or lost focus.
    fn set_focused(&mut self, _focused: bool) {}

    /// Whether the mouse is currently hovering over this widget.
    fn hovered(&self) -> bool { false }

    /// Notify the widget that the mouse entered or left it.
    fn set_hovered(&mut self, _hovered: bool) {}

    /// Unique identifier for this widget instance.
    fn widget_id(&self) -> WidgetId { WidgetId::NONE }

    /// Handle a command sent from the host application.
    /// Returns a `CommandValue` for queries (`GetText`, `GetValue`), or `CommandValue::None`.
    fn handle_command(&mut self, _cmd: &WidgetCommand) -> CommandValue { CommandValue::None }

    /// Get the tooltip text for this widget. Empty string means no tooltip.
    fn tooltip(&self) -> &str { "" }

    /// Set the tooltip text for this widget.
    fn set_tooltip(&mut self, _tooltip: &str) {}
}

// ── NullWidget ─────────────────────────────────────────────────────────

/// A no-op widget used as a placeholder in containers.
pub struct NullWidget {
    rect: LayoutRect,
}

impl NullWidget {
    pub fn new() -> Self { Self { rect: LayoutRect::zero() } }
}

impl PanelWidget for NullWidget {
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; }
    fn rect(&self) -> LayoutRect { self.rect }
    fn render(&mut self, _ctx: &mut RenderContext) {}
    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool { false }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
}

// ── FocusManager ───────────────────────────────────────────────────────

/// Manages keyboard focus across a collection of `PanelWidget`s.
///
/// Tracks which widget index is focused, supports Tab/Shift+Tab cycling,
/// and routes key events to the focused widget.
pub struct FocusManager {
    /// Index of the currently focused widget (into the caller's widget list).
    focused: Option<usize>,
    /// Index of the widget the mouse is currently hovering over.
    hovered: Option<usize>,
    /// Tooltip display state: (hover_start_frame, mouse_x, mouse_y).
    tooltip_state: Option<(u64, f32, f32)>,
    /// Frame counter incremented each time hover is updated.
    frame_counter: u64,
    /// Number of frames to wait before showing a tooltip (approx ~60fps → 45 frames ≈ 750ms).
    tooltip_delay_frames: u64,
}

impl FocusManager {
    pub fn new() -> Self {
        Self { focused: None, hovered: None, tooltip_state: None, frame_counter: 0, tooltip_delay_frames: 45 }
    }

    /// The currently focused widget index.
    pub fn focused(&self) -> Option<usize> { self.focused }

    /// Set the focused widget index directly.
    pub fn set_focused(&mut self, idx: Option<usize>) { self.focused = idx; }

    /// Focus the next focusable widget (Tab key).
    /// `count` is the total number of widgets.
    pub fn focus_next(&mut self, widgets: &mut [Box<dyn PanelWidget>]) {
        let len = widgets.len();
        if len == 0 { return; }
        let start = self.focused.map(|i| i + 1).unwrap_or(0);
        for offset in 0..len {
            let idx = (start + offset) % len;
            if widgets[idx].focusable() {
                self.apply_focus(widgets, Some(idx));
                return;
            }
        }
    }

    /// Focus the previous focusable widget (Shift+Tab key).
    pub fn focus_prev(&mut self, widgets: &mut [Box<dyn PanelWidget>]) {
        let len = widgets.len();
        if len == 0 { return; }
        let start = self.focused.unwrap_or(0).wrapping_sub(1);
        for offset in 0..len {
            let idx = (start.wrapping_sub(offset)) % len;
            if idx < len && widgets[idx].focusable() {
                self.apply_focus(widgets, Some(idx));
                return;
            }
        }
    }

    /// Focus the widget at the given position (e.g. on mouse click).
    pub fn focus_at(&mut self, widgets: &mut [Box<dyn PanelWidget>], x: f32, y: f32) {
        for (i, w) in widgets.iter().enumerate() {
            if w.focusable() && w.rect().contains(x, y) {
                self.apply_focus(widgets, Some(i));
                return;
            }
        }
        // Click on non-focusable area clears focus
        self.apply_focus(widgets, None);
    }

    /// Handle a key event: Tab/Shift+Tab for cycling, otherwise route to focused widget.
    /// Returns `true` if the event was consumed.
    pub fn handle_key(&mut self, widgets: &mut [Box<dyn PanelWidget>], event: &KeyEvent) -> bool {
        use winit::keyboard::Key::Named;
        use winit::keyboard::NamedKey;
        if event.state == winit::event::ElementState::Pressed {
            if let Named(NamedKey::Tab) = &event.logical_key {
                if event.shift {
                    self.focus_prev(widgets);
                } else {
                    self.focus_next(widgets);
                }
                return true;
            }
        }
        // Route to focused widget
        if let Some(idx) = self.focused {
            if idx < widgets.len() {
                return widgets[idx].handle_key(event);
            }
        }
        false
    }

    /// Handle a mouse event: update focus on click, update hover, then delegate.
    /// Returns `true` if the event was consumed.
    pub fn handle_mouse(&mut self, widgets: &mut [Box<dyn PanelWidget>], event: &MouseEvent) -> bool {
        if matches!(event.kind, MouseEventKind::Press(_)) {
            self.focus_at(widgets, event.x, event.y);
        }
        // Update hover tracking on every mouse event
        self.update_hover(widgets, event.x, event.y);
        // Delegate to all widgets (not just focused — for hover etc.)
        let mut consumed = false;
        for w in widgets.iter_mut() {
            if w.handle_mouse(event) {
                consumed = true;
            }
        }
        consumed
    }

    /// Update hover state: fire MouseEnter/MouseLeave when the hovered widget changes.
    pub fn update_hover(&mut self, widgets: &mut [Box<dyn PanelWidget>], x: f32, y: f32) {
        self.frame_counter += 1;
        let mut new_hover: Option<usize> = None;
        // Find topmost widget under cursor (last in list = topmost)
        for (i, w) in widgets.iter().enumerate().rev() {
            if w.rect().contains(x, y) {
                new_hover = Some(i);
                break;
            }
        }
        if new_hover != self.hovered {
            // Leave old
            if let Some(old) = self.hovered {
                if old < widgets.len() {
                    widgets[old].set_hovered(false);
                }
            }
            // Enter new
            if let Some(idx) = new_hover {
                if idx < widgets.len() {
                    widgets[idx].set_hovered(true);
                }
            }
            self.hovered = new_hover;
            // Reset tooltip timer on hover change
            self.tooltip_state = new_hover.map(|_| (self.frame_counter, x, y));
        }
    }

    /// Render all widgets, drawing a focus ring on the focused one.
    pub fn render_all(&self, widgets: &mut [Box<dyn PanelWidget>], ctx: &mut RenderContext) {
        for (i, w) in widgets.iter_mut().enumerate() {
            w.render(ctx);
            if self.focused == Some(i) {
                Self::draw_focus_ring(ctx, w.rect());
            }
        }
        // Render tooltip overlay if applicable
        self.render_tooltip(widgets, ctx);
    }

    /// Render the tooltip for the currently hovered widget after the delay has elapsed.
    fn render_tooltip(&self, widgets: &[Box<dyn PanelWidget>], ctx: &mut RenderContext) {
        let Some(idx) = self.hovered else { return; };
        let Some((start_frame, mx, my)) = self.tooltip_state else { return; };
        if self.frame_counter.saturating_sub(start_frame) < self.tooltip_delay_frames { return; }
        if idx >= widgets.len() { return; }
        let tip = widgets[idx].tooltip();
        if tip.is_empty() { return; }

        let font_size = 12.0_f32;
        let padding = 4.0_f32;
        let text_w = super::ide_text::measure_text(ctx.font_system, tip, font_size, ctx.scale);
        let tip_w = text_w + padding * 2.0;
        let tip_h = font_size + padding * 2.0;
        let tx = mx + 12.0;
        let ty = my + 18.0;

        let scale = ctx.scale;
        let ts = tiny_skia::Transform::from_scale(scale, scale);

        // Background
        if let Some(rect) = tiny_skia::Rect::from_xywh(tx * scale, ty * scale, tip_w * scale, tip_h * scale) {
            let mut bg = tiny_skia::Paint::default();
            bg.set_color_rgba8(255, 255, 225, 240); // light yellow tooltip bg
            ctx.pixmap.fill_rect(rect, &bg, tiny_skia::Transform::identity(), None);
            // Border
            let mut border = tiny_skia::Paint::default();
            border.set_color_rgba8(100, 100, 100, 200);
            let mut stroke = tiny_skia::Stroke::default();
            stroke.width = 1.0;
            if let Some(path) = super::rounded_rect_path(tx, ty, tip_w, tip_h, 2.0) {
                ctx.pixmap.stroke_path(&path, &border, &stroke, ts, None);
            }
        }

        // Text
        super::ide_text::draw_text(
            ctx.pixmap, ctx.font_system, ctx.swash_cache,
            tip, tx + padding, ty + padding - 1.0, font_size,
            cosmic_text::Color::rgba(20, 20, 20, 255), scale,
        );
    }

    /// Drain events from all widgets into a single vec.
    /// Also enforces radio-group mutual exclusion: when a `RadioSelected` event
    /// is seen, other radios in the same group are deselected via commands.
    pub fn drain_all_events(&mut self, widgets: &mut [Box<dyn PanelWidget>]) -> Vec<WidgetEvent> {
        let mut all = Vec::new();
        for w in widgets.iter_mut() {
            all.append(&mut w.drain_events());
        }
        // Enforce radio group mutual exclusion
        for ev in &all {
            if let WidgetEvent::RadioSelected(selected_name, true) = ev {
                // Find the group of the selected radio
                let get_group = WidgetCommand::Custom("GetGroup".into(), CommandValue::None);
                let mut group = String::new();
                for w in widgets.iter_mut() {
                    if w.name() == selected_name.as_str() {
                        if let CommandValue::Text(g) = w.handle_command(&get_group) {
                            group = g;
                        }
                        break;
                    }
                }
                if !group.is_empty() {
                    // Deselect all other radios in the same group
                    for w in widgets.iter_mut() {
                        if w.name() != selected_name.as_str() {
                            if let CommandValue::Text(g) = w.handle_command(&get_group) {
                                if g == group {
                                    w.handle_command(&WidgetCommand::SetChecked(false));
                                }
                            }
                        }
                    }
                }
            }
        }
        all
    }

    /// Send a command to the first widget whose `name()` matches.
    /// Returns the command result, or `CommandValue::None` if no widget matched.
    pub fn send_command(&mut self, widgets: &mut [Box<dyn PanelWidget>], name: &str, cmd: &WidgetCommand) -> CommandValue {
        for w in widgets.iter_mut() {
            if w.name() == name {
                return w.handle_command(cmd);
            }
        }
        CommandValue::None
    }

    /// Send a command to the widget with the given `WidgetId`.
    /// Returns the command result, or `CommandValue::None` if no widget matched.
    pub fn send_command_by_id(&mut self, widgets: &mut [Box<dyn PanelWidget>], id: WidgetId, cmd: &WidgetCommand) -> CommandValue {
        for w in widgets.iter_mut() {
            if w.widget_id() == id {
                return w.handle_command(cmd);
            }
        }
        CommandValue::None
    }

    /// Broadcast a command to all widgets.
    /// Returns a `Vec` of `(WidgetId, CommandValue)` from widgets that returned a non-None value.
    pub fn broadcast_command(&mut self, widgets: &mut [Box<dyn PanelWidget>], cmd: &WidgetCommand) -> Vec<(WidgetId, CommandValue)> {
        let mut results = Vec::new();
        for w in widgets.iter_mut() {
            let val = w.handle_command(cmd);
            if !matches!(val, CommandValue::None) {
                results.push((w.widget_id(), val));
            }
        }
        results
    }

    // ── internal ──

    fn apply_focus(&mut self, widgets: &mut [Box<dyn PanelWidget>], new: Option<usize>) {
        if self.focused == new { return; }
        if let Some(old) = self.focused {
            if old < widgets.len() {
                widgets[old].set_focused(false);
            }
        }
        self.focused = new;
        if let Some(idx) = new {
            if idx < widgets.len() {
                widgets[idx].set_focused(true);
            }
        }
    }

    fn draw_focus_ring(ctx: &mut RenderContext, r: LayoutRect) {
        let scale = ctx.scale;
        let px = (r.x * scale) as i32;
        let py = (r.y * scale) as i32;
        let pw = (r.w * scale) as i32;
        let ph = (r.h * scale) as i32;
        let color = ColorU8::from_rgba(0, 120, 215, 200).premultiply();
        let thickness = (2.0 * scale) as i32;
        let pix = &mut *ctx.pixmap;
        let w = pix.width() as i32;
        let h = pix.height() as i32;
        // top edge
        for dy in 0..thickness {
            let y = py + dy;
            if y < 0 || y >= h { continue; }
            for x in px..(px + pw).min(w) {
                if x >= 0 {
                    pix.pixels_mut()[(y * w + x) as usize] = color;
                }
            }
        }
        // bottom edge
        for dy in 0..thickness {
            let y = py + ph - 1 - dy;
            if y < 0 || y >= h { continue; }
            for x in px..(px + pw).min(w) {
                if x >= 0 {
                    pix.pixels_mut()[(y * w + x) as usize] = color;
                }
            }
        }
        // left edge
        for dx in 0..thickness {
            let x = px + dx;
            if x < 0 || x >= w { continue; }
            for y in py..(py + ph).min(h) {
                if y >= 0 {
                    pix.pixels_mut()[(y * w + x) as usize] = color;
                }
            }
        }
        // right edge
        for dx in 0..thickness {
            let x = px + pw - 1 - dx;
            if x < 0 || x >= w { continue; }
            for y in py..(py + ph).min(h) {
                if y >= 0 {
                    pix.pixels_mut()[(y * w + x) as usize] = color;
                }
            }
        }
    }
}
