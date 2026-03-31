use tiny_skia::*;
use super::WidgetColors;

pub struct ToolStrip {
    pub items: Vec<String>,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl ToolStrip {
    pub fn new() -> Self { Self { items: Vec::new(), width: 400.0, height: 28.0, colors: WidgetColors::default() } }
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default(); paint.anti_alias = true;
        paint.set_color_rgba8(245,245,245,255);
        if let Some(r) = Rect::from_xywh(x,y,self.width,self.height) { pixmap.fill_rect(r, &paint, ts, None); }
    }
    pub fn measure(&self)->(f32,f32){(self.width,self.height)}
}
