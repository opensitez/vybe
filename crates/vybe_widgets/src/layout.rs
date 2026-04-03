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
    /// Generic named action (for custom widgets).
    Action(String),
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
