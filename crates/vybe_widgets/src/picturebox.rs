use tiny_skia::*;
use super::WidgetColors;

pub struct PictureBox {
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl PictureBox {
    pub fn new() -> Self { Self { width: 160.0, height: 120.0, colors: WidgetColors::default() } }
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default(); paint.anti_alias = true;
        paint.set_color_rgba8(240,240,240,255);
        if let Some(r) = Rect::from_xywh(x,y,self.width,self.height) { pixmap.fill_rect(r, &paint, ts, None); }
        // Placeholder text is left to caller
        paint.set_color_rgba8(200,200,200,255);
        let mut stroke = Stroke::default(); stroke.width = 1.0;
        let mut pb = PathBuilder::new(); pb.move_to(x,y); pb.line_to(x+self.width,y); pb.line_to(x+self.width,y+self.height); pb.line_to(x,y+self.height); pb.close();
        if let Some(path) = pb.finish() { pixmap.stroke_path(&path, &paint, &stroke, ts, None); }
    }
    pub fn measure(&self)->(f32,f32){(self.width,self.height)}
}
