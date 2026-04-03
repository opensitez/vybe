//! Checkbox widget — standalone tiny-skia rendered checkbox.

use tiny_skia::*;
use cosmic_text::Color as CosmicColor;
use super::{WidgetColors, rounded_rect_path};
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent, WidgetId};

pub struct Checkbox {
    pub checked: bool,
    pub label: String,
    pub disabled: bool,
    pub focused: bool,
    pub hovered: bool,
    pub colors: WidgetColors,
    pub size: f32,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl Checkbox {
    pub fn new(label: &str) -> Self {
        Self {
            checked: false,
            label: label.to_string(),
            disabled: false,
            focused: false,
            hovered: false,
            colors: WidgetColors::default(),
            size: 16.0,
            id: WidgetId::next(),
            name: label.to_string(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Paint the checkbox at (x, y) into the pixmap.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let sz = self.size;
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(path) = rounded_rect_path(x, y, sz, sz, 2.0) {
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
        }

        // Border
        let (r, g, b, a) = if self.focused { self.colors.focus_ring } else { self.colors.border };
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = if self.focused { 2.0 } else { 1.0 };
        if let Some(path) = rounded_rect_path(x, y, sz, sz, 2.0) {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Check mark
        if self.checked {
            let cx = x + sz / 2.0;
            let cy = y + sz / 2.0;
            let s = sz * 0.3;
            let (r, g, b, a) = self.colors.foreground;
            paint.set_color_rgba8(r, g, b, a);
            stroke.width = 2.0;
            let mut pb = PathBuilder::new();
            pb.move_to(cx - s, cy);
            pb.line_to(cx - s * 0.3, cy + s * 0.7);
            pb.line_to(cx + s, cy - s * 0.6);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }
    }

    /// Returns (width, height) for layout.
    pub fn measure(&self) -> (f32, f32) {
        (self.size, self.size)
    }

    /// Handle a click at (x, y) relative to the widget's origin.
    /// Returns true if the checkbox was toggled.
    pub fn click(&mut self, x: f32, y: f32) -> bool {
        if self.disabled { return false; }
        if x >= 0.0 && y >= 0.0 && x <= self.size && y <= self.size {
            self.checked = !self.checked;
            return true;
        }
        false
    }

    /// Toggle the checked state.
    pub fn toggle(&mut self) {
        if !self.disabled { self.checked = !self.checked; }
    }
}

// ── PanelWidget impl ───────────────────────────────────────────────────

impl PanelWidget for Checkbox {
    fn name(&self) -> &str { &self.name }
    fn widget_id(&self) -> WidgetId { self.id }
    fn set_focused(&mut self, focused: bool) { self.focused = focused; }
    fn hovered(&self) -> bool { self.hovered }
    fn set_hovered(&mut self, hovered: bool) { self.hovered = hovered; }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
    }

    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }

        // Render checkbox box via existing paint method
        let box_y = r.y + (r.h - self.size) / 2.0;
        self.paint(ctx.pixmap, r.x, box_y, ctx.scale);

        // Render label text next to checkbox
        if !self.label.is_empty() {
            let (fr, fg, fb, _) = self.colors.foreground;
            let font_size = 13.0;
            let tx = r.x + self.size + 6.0;
            let ty = r.y + (r.h - font_size) / 2.0 - 1.0;
            super::ide_text::draw_text(
                ctx.pixmap, ctx.font_system, ctx.swash_cache,
                &self.label, tx, ty, font_size,
                CosmicColor::rgba(fr, fg, fb, 255), ctx.scale,
            );
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) { return false; }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            if !self.disabled {
                self.checked = !self.checked;
                self.pending_events.push(WidgetEvent::CheckboxToggled(self.name.clone(), self.checked));
            }
            return true;
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if !self.focused { return false; }
        use winit::keyboard::{Key, NamedKey};
        use winit::event::ElementState;
        if event.state == ElementState::Pressed {
            if let Key::Named(NamedKey::Space) = &event.key_without_modifiers {
                if !self.disabled {
                    self.checked = !self.checked;
                    self.pending_events.push(WidgetEvent::CheckboxToggled(self.name.clone(), self.checked));
                }
                return true;
            }
        }
        false
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }

    fn focusable(&self) -> bool { !self.disabled }
}
