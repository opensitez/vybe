//! TableLayoutPanel widget — grid with dotted cell borders.

use tiny_skia::*;
use super::WidgetColors;
use super::layout::{LayoutRect, MouseEvent, KeyEvent, RenderContext, PanelWidget, WidgetEvent};

pub struct TableLayoutPanel {
    pub cols: usize,
    pub rows: usize,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub name: String,
    rect: LayoutRect,
}

impl TableLayoutPanel {
    pub fn new(cols: usize, rows: usize) -> Self {
        Self {
            cols,
            rows,
            width: 300.0,
            height: 200.0,
            colors: WidgetColors {
                background: (255, 255, 255, 255),
                border: (200, 200, 200, 255),
                ..WidgetColors::default()
            },
            name: String::new(),
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Cell dimensions.
    pub fn cell_size(&self) -> (f32, f32) {
        let cw = self.width / self.cols.max(1) as f32;
        let ch = self.height / self.rows.max(1) as f32;
        (cw, ch)
    }

    /// Get rect for a specific cell (col, row).
    pub fn cell_rect(&self, col: usize, row: usize) -> (f32, f32, f32, f32) {
        let (cw, ch) = self.cell_size();
        (col as f32 * cw, row as f32 * ch, cw, ch)
    }

    /// Paint — white background with dotted grid lines.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // White background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Dotted cell borders
        let (r, g, b, a) = self.colors.border;
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        stroke.dash = StrokeDash::new(vec![2.0, 2.0], 0.0);

        let (cw, ch) = self.cell_size();

        // Vertical lines
        for i in 0..=self.cols {
            let cx = x + i as f32 * cw;
            let mut pb = PathBuilder::new();
            pb.move_to(cx, y);
            pb.line_to(cx, y + self.height);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }

        // Horizontal lines
        for j in 0..=self.rows {
            let ry = y + j as f32 * ch;
            let mut pb = PathBuilder::new();
            pb.move_to(x, ry);
            pb.line_to(x + self.width, ry);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }

        // Solid outer border
        let mut solid_stroke = Stroke::default();
        solid_stroke.width = 1.0;
        paint.set_color_rgba8(180, 180, 180, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &solid_stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}

impl PanelWidget for TableLayoutPanel {
    fn name(&self) -> &str { &self.name }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; self.width = rect.w; self.height = rect.h; }
    fn rect(&self) -> LayoutRect { self.rect }
    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
    }
    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool { false }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
}
