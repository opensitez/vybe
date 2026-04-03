//! Radio button widget — standalone tiny-skia rendered radio button.

use tiny_skia::*;
use cosmic_text::Color as CosmicColor;
use super::{WidgetColors, circle_path};
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent};

pub struct Radio {
    pub selected: bool,
    pub label: String,
    pub disabled: bool,
    pub focused: bool,
    pub hovered: bool,
    pub colors: WidgetColors,
    pub size: f32,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl Radio {
    pub fn new(label: &str) -> Self {
        Self {
            selected: false,
            label: label.to_string(),
            disabled: false,
            focused: false,
            hovered: false,
            colors: WidgetColors::default(),
            size: 16.0,
            name: label.to_string(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let sz = self.size;
        let cx = x + sz / 2.0;
        let cy = y + sz / 2.0;
        let r = sz / 2.0;
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Outer circle
        if let Some(path) = circle_path(cx, cy, r) {
            let (cr, cg, cb, ca) = self.colors.background;
            paint.set_color_rgba8(cr, cg, cb, ca);
            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);

            let (cr, cg, cb, ca) = if self.focused { self.colors.focus_ring } else { self.colors.border };
            paint.set_color_rgba8(cr, cg, cb, ca);
            let mut stroke = Stroke::default();
            stroke.width = if self.focused { 2.0 } else { 1.0 };
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Inner dot
        if self.selected {
            let ir = r * 0.45;
            if let Some(path) = circle_path(cx, cy, ir) {
                let (cr, cg, cb, ca) = self.colors.foreground;
                paint.set_color_rgba8(cr, cg, cb, ca);
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.size, self.size)
    }

    pub fn click(&mut self, x: f32, y: f32) -> bool {
        if self.disabled { return false; }
        let cx = self.size / 2.0;
        let cy = self.size / 2.0;
        let dist = ((x - cx).powi(2) + (y - cy).powi(2)).sqrt();
        if dist <= self.size / 2.0 {
            self.selected = true;
            return true;
        }
        false
    }
}

impl PanelWidget for Radio {
    fn name(&self) -> &str { &self.name }
    fn set_focused(&mut self, focused: bool) { self.focused = focused; }
    fn hovered(&self) -> bool { self.hovered }
    fn set_hovered(&mut self, hovered: bool) { self.hovered = hovered; }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; }
    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        let box_y = r.y + (r.h - self.size) / 2.0;
        self.paint(ctx.pixmap, r.x, box_y, ctx.scale);
        if !self.label.is_empty() {
            let (fr, fg, fb, _) = self.colors.foreground;
            let tx = r.x + self.size + 6.0;
            let ty = r.y + (r.h - 13.0) / 2.0 - 1.0;
            super::ide_text::draw_text(ctx.pixmap, ctx.font_system, ctx.swash_cache, &self.label, tx, ty, 13.0, CosmicColor::rgba(fr, fg, fb, 255), ctx.scale);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) { return false; }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            if !self.disabled {
                self.selected = true;
                self.pending_events.push(WidgetEvent::RadioSelected(self.name.clone(), true));
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
                    self.selected = true;
                    self.pending_events.push(WidgetEvent::RadioSelected(self.name.clone(), true));
                }
                return true;
            }
        }
        false
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> { std::mem::take(&mut self.pending_events) }
    fn focusable(&self) -> bool { !self.disabled }
}
