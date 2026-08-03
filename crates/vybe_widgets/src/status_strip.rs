//! StatusStrip widget — colored status bar at bottom of window.

use super::WidgetColors;
use super::layout::{
    KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent, MouseEventKind,
    PanelWidget, RenderContext, WidgetEvent, WidgetId };
use cosmic_text::Color as CosmicColor;
use tiny_skia::*;

pub struct StatusStrip {
    pub text: String,
    pub items: Vec<String>,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent> }

impl StatusStrip {
    pub fn new() -> Self {
        Self {
            text: String::new(),
            items: Vec::new(),
            width: 800.0,
            height: 24.0,
            colors: WidgetColors {
                background: (0, 122, 204, 255), // VS Code blue
                foreground: (255, 255, 255, 255),
                ..WidgetColors::default()
            },
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new() }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Paint — colored bar with top border. Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Top border (slightly darker)
        let dr = r.saturating_sub(20);
        let dg = g.saturating_sub(20);
        let db = b.saturating_sub(20);
        paint.set_color_rgba8(dr, dg, db, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + self.width, y);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Item separator lines (if multiple items)
        if self.items.len() > 1 {
            // Lighter separator
            paint.set_color_rgba8(r.min(235) + 20, g.min(235) + 20, b.min(235) + 20, 100);
            let item_w = self.width / self.items.len() as f32;
            for i in 1..self.items.len() {
                let sx = x + i as f32 * item_w;
                let mut pb = PathBuilder::new();
                pb.move_to(sx, y + 4.0);
                pb.line_to(sx, y + self.height - 4.0);
                if let Some(path) = pb.finish() {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

impl PanelWidget for StatusStrip {
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
        let (fr, fg, fb, _) = self.colors.foreground;
        let col = CosmicColor::rgba(fr, fg, fb, 255);
        if !self.text.is_empty() {
            super::ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &self.text,
                r.x + 6.0,
                r.y + 4.0,
                12.0,
                col,
                ctx.scale,
            );
        } else if !self.items.is_empty() {
            let item_w = r.w / self.items.len() as f32;
            for (i, item) in self.items.iter().enumerate() {
                let ix = r.x + i as f32 * item_w + 6.0;
                super::ide_text::draw_text(
                    ctx.pixmap,
                    ctx.font_system,
                    ctx.swash_cache,
                    item,
                    ix,
                    r.y + 4.0,
                    12.0,
                    col,
                    ctx.scale,
                );
            }
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            return false;
        }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            self.pending_events
                .push(WidgetEvent::StatusBarClick(self.name.clone()));
            return true;
        }
        false
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }
    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }
}
