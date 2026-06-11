//! GroupBox widget — border with title gap at top.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetCommand,
    WidgetId,
};
use cosmic_text::Color as CosmicColor;
use tiny_skia::*;

pub struct GroupBox {
    pub title: String,
    pub title_width: f32,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
}

impl GroupBox {
    pub fn new<S: Into<String>>(title: S) -> Self {
        let t: String = title.into();
        Self {
            id: WidgetId::next(),
            name: t.clone(),
            title: t,
            title_width: 60.0,
            width: 200.0,
            height: 120.0,
            colors: WidgetColors {
                border: (160, 160, 160, 255),
                ..WidgetColors::default()
            },
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Paint the groupbox — border with gap at top for title. Title text drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        // The border starts slightly below top to leave room for title text center
        let border_top = y + 8.0;
        let title_gap_start = x + 8.0;
        let title_gap_end = x + 8.0 + self.title_width + 8.0;
        let (r, g, b, a) = self.colors.border;
        paint.set_color_rgba8(r, g, b, a);

        // Top border with gap for title
        // Left part of top border
        let mut pb = PathBuilder::new();
        pb.move_to(x, border_top);
        pb.line_to(title_gap_start, border_top);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Right part of top border (after title gap)
        let mut pb = PathBuilder::new();
        pb.move_to(title_gap_end, border_top);
        pb.line_to(x + self.width, border_top);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Right border
        let mut pb = PathBuilder::new();
        pb.move_to(x + self.width, border_top);
        pb.line_to(x + self.width, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Bottom border
        let mut pb = PathBuilder::new();
        pb.move_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        // Left border
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x, border_top);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Content area offset (children should be placed inside the border).
    pub fn content_rect(&self) -> (f32, f32, f32, f32) {
        (4.0, 18.0, self.width - 8.0, self.height - 22.0)
    }
}

impl PanelWidget for GroupBox {
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
        // Measure title for gap
        self.title_width =
            super::ide_text::measure_text(ctx.font_system, &self.title, 12.0, ctx.scale);
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
        // Draw title text
        let (fr, fg, fb, _) = self.colors.foreground;
        super::ide_text::draw_text(
            ctx.pixmap,
            ctx.font_system,
            ctx.swash_cache,
            &self.title,
            r.x + 12.0,
            r.y,
            12.0,
            CosmicColor::rgba(fr, fg, fb, 255),
            ctx.scale,
        );
    }
    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool {
        false
    }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetText(t) => {
                self.title = t.clone();
                CommandValue::None
            }
            WidgetCommand::GetText => CommandValue::Text(self.title.clone()),
            _ => CommandValue::None,
        }
    }
}
