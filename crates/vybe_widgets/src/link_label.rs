use tiny_skia::*;
use super::WidgetColors;

pub struct LinkLabel {
    pub text: String,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl LinkLabel {
    pub fn new<S: Into<String>>(text: S) -> Self { Self { text: text.into(), width: 100.0, height: 20.0, colors: WidgetColors::default() } }
    pub fn paint(&self, _pixmap: &mut Pixmap, _x: f32, _y: f32, _scale: f32) {}
    pub fn measure(&self)->(f32,f32){(self.width,self.height)}
}
