//! TableLayoutPanel — arranges children in a grid of rows × columns.
//!
//! Children are assigned to cells via (col, row) coordinates. Each cell can hold
//! one widget. Column widths and row heights can be fixed or proportional.

use tiny_skia::*;
use super::WidgetColors;
use super::layout::{LayoutRect, MouseEvent, KeyEvent, RenderContext, PanelWidget, WidgetEvent, WidgetId, WidgetCommand, CommandValue};

/// How a column or row is sized.
#[derive(Clone, Copy, Debug)]
pub enum SizeMode {
    /// Fixed pixel size.
    Absolute(f32),
    /// Proportional share of remaining space (weight).
    Percent(f32),
}

/// A child placed in a specific cell.
struct CellChild {
    col: usize,
    row: usize,
    col_span: usize,
    row_span: usize,
    widget: Box<dyn PanelWidget>,
}

pub struct TableLayoutPanel {
    pub cols: usize,
    pub rows: usize,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    /// Padding inside each cell.
    pub cell_padding: f32,
    col_sizes: Vec<SizeMode>,
    row_sizes: Vec<SizeMode>,
    children: Vec<CellChild>,
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
            id: WidgetId::next(),
            name: String::new(),
            cell_padding: 2.0,
            col_sizes: vec![SizeMode::Percent(1.0); cols],
            row_sizes: vec![SizeMode::Percent(1.0); rows],
            children: Vec::new(),
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Set column sizing. `sizes` length should match `cols`.
    pub fn set_col_sizes(&mut self, sizes: Vec<SizeMode>) {
        self.col_sizes = sizes;
        // Pad or truncate to match cols
        self.col_sizes.resize(self.cols, SizeMode::Percent(1.0));
        self.relayout();
    }

    /// Set row sizing. `sizes` length should match `rows`.
    pub fn set_row_sizes(&mut self, sizes: Vec<SizeMode>) {
        self.row_sizes = sizes;
        self.row_sizes.resize(self.rows, SizeMode::Percent(1.0));
        self.relayout();
    }

    /// Add a child to a specific cell.
    pub fn add(&mut self, col: usize, row: usize, widget: Box<dyn PanelWidget>) {
        self.children.push(CellChild { col, row, col_span: 1, row_span: 1, widget });
        self.relayout();
    }

    /// Add a child spanning multiple cells.
    pub fn add_span(&mut self, col: usize, row: usize, col_span: usize, row_span: usize, widget: Box<dyn PanelWidget>) {
        self.children.push(CellChild { col, row, col_span: col_span.max(1), row_span: row_span.max(1), widget });
        self.relayout();
    }

    pub fn child_count(&self) -> usize { self.children.len() }

    /// Resolve sizes into pixel offsets. Returns cumulative offsets (len = count+1).
    fn resolve_sizes(modes: &[SizeMode], total: f32) -> Vec<f32> {
        let mut offsets = Vec::with_capacity(modes.len() + 1);
        // First pass: sum absolute and percent weights
        let mut abs_total = 0.0_f32;
        let mut pct_total = 0.0_f32;
        for m in modes {
            match m {
                SizeMode::Absolute(px) => abs_total += px,
                SizeMode::Percent(w) => pct_total += w,
            }
        }
        let remaining = (total - abs_total).max(0.0);
        let mut pos = 0.0;
        offsets.push(pos);
        for m in modes {
            let size = match m {
                SizeMode::Absolute(px) => *px,
                SizeMode::Percent(w) => {
                    if pct_total > 0.0 { remaining * w / pct_total } else { 0.0 }
                }
            };
            pos += size;
            offsets.push(pos);
        }
        offsets
    }

    fn relayout(&mut self) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }

        let col_offsets = Self::resolve_sizes(&self.col_sizes, r.w);
        let row_offsets = Self::resolve_sizes(&self.row_sizes, r.h);
        let pad = self.cell_padding;

        for child in &mut self.children {
            let c = child.col.min(self.cols.saturating_sub(1));
            let rv = child.row.min(self.rows.saturating_sub(1));
            let c_end = (c + child.col_span).min(self.cols);
            let r_end = (rv + child.row_span).min(self.rows);

            let x0 = col_offsets.get(c).copied().unwrap_or(0.0);
            let x1 = col_offsets.get(c_end).copied().unwrap_or(r.w);
            let y0 = row_offsets.get(rv).copied().unwrap_or(0.0);
            let y1 = row_offsets.get(r_end).copied().unwrap_or(r.h);

            let cell_rect = LayoutRect::new(
                r.x + x0 + pad,
                r.y + y0 + pad,
                (x1 - x0 - pad * 2.0).max(0.0),
                (y1 - y0 - pad * 2.0).max(0.0),
            );
            child.widget.set_rect(cell_rect);
        }
    }

    /// Cell dimensions (uniform sizing, for backward compat).
    pub fn cell_size(&self) -> (f32, f32) {
        let cw = self.width / self.cols.max(1) as f32;
        let ch = self.height / self.rows.max(1) as f32;
        (cw, ch)
    }

    /// Get rect for a specific cell (col, row) — uniform sizing.
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

        // Dotted cell borders using resolved offsets
        let (r, g, b, a) = self.colors.border;
        paint.set_color_rgba8(r, g, b, a);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        stroke.dash = StrokeDash::new(vec![2.0, 2.0], 0.0);

        let col_offsets = Self::resolve_sizes(&self.col_sizes, self.width);
        let row_offsets = Self::resolve_sizes(&self.row_sizes, self.height);

        for &cx_off in &col_offsets {
            let cx = x + cx_off;
            let mut pb = PathBuilder::new();
            pb.move_to(cx, y);
            pb.line_to(cx, y + self.height);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
        }

        for &ry_off in &row_offsets {
            let ry = y + ry_off;
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
    fn widget_id(&self) -> WidgetId { self.id }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.width = rect.w;
        self.height = rect.h;
        self.relayout();
    }
    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
        for child in &mut self.children {
            child.widget.render(ctx);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        for child in self.children.iter_mut().rev() {
            if child.widget.rect().contains(event.x, event.y) && child.widget.handle_mouse(event) {
                return true;
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        for child in &mut self.children {
            if child.widget.handle_key(event) {
                return true;
            }
        }
        false
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        for child in self.children.iter_mut().rev() {
            if child.widget.rect().contains(x, y) && child.widget.handle_scroll(delta, x, y) {
                return true;
            }
        }
        false
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        for child in self.children.iter().rev() {
            if child.widget.rect().contains(x, y) {
                return child.widget.cursor_at(x, y);
            }
        }
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = Vec::new();
        for child in &mut self.children {
            events.extend(child.widget.drain_events());
        }
        events
    }

    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetEnabled(_) | WidgetCommand::SetVisible(_) => {
                for child in &mut self.children {
                    child.widget.handle_command(cmd);
                }
                CommandValue::None
            }
            _ => CommandValue::None,
        }
    }
}
