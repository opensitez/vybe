//! Simple data grid widget — header + rows grid rendering.

use super::WidgetColors;
use super::layout::{
    CommandValue, KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetCommand,
    WidgetId,
};
use tiny_skia::*;

pub struct DataGrid {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<String>>,
    pub width: f32,
    pub height: f32,
    pub row_height: f32,
    pub header_height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    /// The headers' and cells' text style — one spec, because the header is
    /// not a separate element here. It was a literal `12.0` at both draw calls,
    /// and `DataGrid` had no `handle_command` at all, so every declaration a
    /// program made about a `<table>`'s text was answered by the trait default
    /// and discarded.
    pub font: crate::ide_text::FontSpec,
    rect: LayoutRect,
}

impl DataGrid {
    pub fn new(cols: &[&str]) -> Self {
        Self {
            columns: cols.iter().map(|s| s.to_string()).collect(),
            rows: Vec::new(),
            width: 400.0,
            height: 200.0,
            row_height: 20.0,
            header_height: 24.0,
            colors: WidgetColors::default(),
            id: WidgetId::next(),
            name: String::new(),
            font: crate::ide_text::FontSpec::sans(12.0),
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        paint.set_color_rgba8(255, 255, 255, 255);
        let rect = Rect::from_xywh(x, y, self.width, self.height).unwrap();
        pixmap.fill_rect(rect, &paint, ts, None);

        // Header background
        paint.set_color_rgba8(245, 245, 245, 255);
        let header_rect = Rect::from_xywh(x, y, self.width, self.header_height).unwrap();
        pixmap.fill_rect(header_rect, &paint, ts, None);

        // Vertical column separators and header text omitted (text handled by external font engine)
        let col_w = if self.columns.is_empty() {
            self.width
        } else {
            self.width / self.columns.len() as f32
        };
        // Draw vertical lines
        paint.set_color_rgba8(200, 200, 200, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        for i in 0..=self.columns.len() {
            let cx = x + i as f32 * col_w;
            let mut pb = PathBuilder::new();
            pb.move_to(cx, y);
            pb.line_to(cx, y + self.height);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }

        // Horizontal lines for header + rows
        let mut row_y = y + self.header_height;
        let mut line_count = 0;
        while row_y <= y + self.height + 0.01 {
            let mut pb = PathBuilder::new();
            pb.move_to(x, row_y);
            pb.line_to(x + self.width, row_y);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
            row_y += self.row_height;
            line_count += 1;
            if line_count > 1000 {
                break;
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

impl PanelWidget for DataGrid {
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
        // Draw column headers and row data
        let (fr, fg, fb, _) = self.colors.foreground;
        let col = cosmic_text::Color::rgba(fr, fg, fb, 255);
        let cw = if self.columns.is_empty() {
            self.width
        } else {
            self.width / self.columns.len() as f32
        };
        for (i, header) in self.columns.iter().enumerate() {
            super::ide_text::draw_text_spec(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                header,
                r.x + i as f32 * cw + 4.0,
                r.y + 4.0,
                &self.font,
                col,
                ctx.scale,
            );
        }
        for (ri, row) in self.rows.iter().enumerate() {
            let ry = r.y + self.header_height + ri as f32 * self.row_height;
            if ry > r.y + r.h {
                break;
            }
            for (ci, cell) in row.iter().enumerate() {
                super::ide_text::draw_text_spec(
                    ctx.pixmap,
                    ctx.font_system,
                    ctx.swash_cache,
                    cell,
                    r.x + ci as f32 * cw + 4.0,
                    ry + 2.0,
                    &self.font,
                    col,
                    ctx.scale,
                );
            }
        }
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::Custom(key, val) if self.font.apply_command(key, val) => {
                CommandValue::None
            }
            _ => CommandValue::None,
        }
    }

    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool {
        false
    }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }
}
