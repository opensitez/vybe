//! Label widget — standalone tiny-skia rendered label.

use tiny_skia::*;
use cosmic_text::Color as CosmicColor;
use super::WidgetColors;
use super::layout::{LayoutRect, MouseEvent, KeyEvent, RenderContext, PanelWidget, WidgetId};

pub struct Label {
    pub id: WidgetId,
    pub text: String,
    pub width: f32,
    pub height: f32,
    pub auto_size: bool,
    pub transparent: bool,
    pub colors: WidgetColors,
    pub font_size: f32,
    rect: LayoutRect,
}

impl Label {
    pub fn new<S: Into<String>>(text: S) -> Self {
        Self {
            id: WidgetId::next(),
            text: text.into(),
            width: 100.0,
            height: 20.0,
            auto_size: true,
            transparent: true,
            colors: WidgetColors {
                background: (240, 240, 240, 255),
                ..WidgetColors::default()
            },
            font_size: 13.0,
            rect: LayoutRect::zero(),
        }
    }

    /// Paint the label background (if not transparent). Text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        if self.transparent {
            return;
        }
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background fill
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

// ── PanelWidget impl ───────────────────────────────────────────────────

impl PanelWidget for Label {
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.width = rect.w;
        self.height = rect.h;
    }

    fn rect(&self) -> LayoutRect { self.rect }
    fn widget_id(&self) -> WidgetId { self.id }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }

        // Background (if not transparent)
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);

        // Text
        let (fr, fg, fb, _) = self.colors.foreground;
        let ty = r.y + (r.h - self.font_size) / 2.0 - 1.0;
        super::ide_text::draw_text(
            ctx.pixmap, ctx.font_system, ctx.swash_cache,
            &self.text, r.x + 2.0, ty, self.font_size,
            CosmicColor::rgba(fr, fg, fb, 255), ctx.scale,
        );
    }

    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool { false }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
}
