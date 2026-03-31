use tiny_skia::*;
use super::WidgetColors;

pub struct TableLayoutPanel {
    pub cols: usize,
    pub rows: usize,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl TableLayoutPanel {
    pub fn new(cols: usize, rows: usize) -> Self { Self { cols, rows, width: 300.0, height: 200.0, colors: WidgetColors::default() } }
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default(); paint.anti_alias = true;
        paint.set_color_rgba8(255,255,255,255);
        if let Some(r) = Rect::from_xywh(x,y,self.width,self.height) { pixmap.fill_rect(r, &paint, ts, None); }
        paint.set_color_rgba8(220,220,220,255);
        let mut stroke = Stroke::default(); stroke.width = 1.0;
        let col_w = self.width / self.cols.max(1) as f32;
        let row_h = self.height / self.rows.max(1) as f32;
        for i in 0..=self.cols {
            let cx = x + i as f32 * col_w;
            let mut pb = PathBuilder::new(); pb.move_to(cx,y); pb.line_to(cx,y+self.height);
            if let Some(path) = pb.finish() { pixmap.stroke_path(&path, &paint, &stroke, ts, None); }
        }
        for j in 0..=self.rows {
            let ry = y + j as f32 * row_h;
            let mut pb = PathBuilder::new(); pb.move_to(x,ry); pb.line_to(x+self.width,ry);
            if let Some(path) = pb.finish() { pixmap.stroke_path(&path, &paint, &stroke, ts, None); }
        }
    }
    pub fn measure(&self)->(f32,f32){(self.width,self.height)}
}
