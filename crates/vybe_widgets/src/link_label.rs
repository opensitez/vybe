//! LinkLabel widget — label with underline, like a hyperlink.

use tiny_skia::*;
use cosmic_text::Color as CosmicColor;
use super::WidgetColors;
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent, WidgetId};

pub struct LinkLabel {
    pub text: String,
    pub width: f32,
    pub height: f32,
    pub hovered: bool,
    pub visited: bool,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl LinkLabel {
    pub fn new<S: Into<String>>(text: S) -> Self {
        let t: String = text.into();
        Self {
            id: WidgetId::next(),
            name: t.clone(),
            text: t,
            width: 100.0,
            height: 20.0,
            hovered: false,
            visited: false,
            colors: WidgetColors {
                foreground: (0, 102, 204, 255),
                background: (0, 0, 0, 0),
                ..WidgetColors::default()
            },
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Paint the link label — underline beneath text area. Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Underline path beneath text area
        let underline_y = y + self.height - 2.0;
        let link_color = if self.visited {
            (128, 0, 128, 255) // purple for visited
        } else if self.hovered {
            (0, 70, 180, 255) // darker blue on hover
        } else {
            self.colors.foreground
        };
        let (r, g, b, a) = link_color;
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        let mut pb = PathBuilder::new();
        pb.move_to(x, underline_y);
        pb.line_to(x + self.width, underline_y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Handle click — returns true if within bounds.
    pub fn click(&self, x: f32, y: f32) -> bool {
        x >= 0.0 && y >= 0.0 && x <= self.width && y <= self.height
    }
}

impl PanelWidget for LinkLabel {
    fn name(&self) -> &str { &self.name }
    fn widget_id(&self) -> WidgetId { self.id }
    fn hovered(&self) -> bool { self.hovered }
    fn set_hovered(&mut self, hovered: bool) { self.hovered = hovered; }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; self.width = rect.w; self.height = rect.h; }
    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
        let link_color = if self.visited { (128, 0, 128, 255) } else if self.hovered { (0, 70, 180, 255) } else { self.colors.foreground };
        let (cr, cg, cb, _) = link_color;
        super::ide_text::draw_text(ctx.pixmap, ctx.font_system, ctx.swash_cache, &self.text, r.x, r.y + 1.0, 13.0, CosmicColor::rgba(cr, cg, cb, 255), ctx.scale);
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            self.hovered = false;
            return false;
        }
        self.hovered = true;
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            self.visited = true;
            self.pending_events.push(WidgetEvent::LinkClicked(self.name.clone()));
            return true;
        }
        false
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
    fn drain_events(&mut self) -> Vec<WidgetEvent> { std::mem::take(&mut self.pending_events) }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        if self.rect.contains(x, y) { winit::window::CursorIcon::Pointer } else { winit::window::CursorIcon::Default }
    }
}
