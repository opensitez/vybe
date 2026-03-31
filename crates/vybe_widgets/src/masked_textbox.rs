use tiny_skia::*;
use super::WidgetColors;

pub struct MaskedTextBox {
    pub text: String,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl MaskedTextBox {
    pub fn new() -> Self { Self { text: String::new(), width: 140.0, height: 24.0, colors: WidgetColors::default() } }
    pub fn paint(&self, _pixmap: &mut Pixmap, _x: f32, _y: f32, _scale: f32) {}
    pub fn measure(&self)->(f32,f32){(self.width,self.height)}
}
