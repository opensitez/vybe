//! Layout primitives and the `PanelWidget` trait.
//!
//! This module provides the foundation for the vybe GUI toolkit:
//! - `LayoutRect` — rectangle with hit-testing and subdivision helpers
//! - `MouseEvent` / `KeyEvent` — unified input events
//! - `RenderContext` — bundle of rendering resources
//! - `PanelWidget` trait — implemented by all toolkit panels and containers

use cosmic_text::{FontSystem, SwashCache};
use tiny_skia::Pixmap;

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

    /// Whether this widget wants keyboard focus.
    fn focusable(&self) -> bool { false }
}
