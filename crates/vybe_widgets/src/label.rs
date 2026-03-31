use tiny_skia::*;
use super::WidgetColors;

pub struct Label {
    pub text: String,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl Label {
    pub fn new<S: Into<String>>(text: S) -> Self {
        Self { text: text.into(), width: 100.0, height: 20.0, colors: WidgetColors::default() }
    }

    pub fn paint(&self, _pixmap: &mut Pixmap, _x: f32, _y: f32, _scale: f32) {
        // Text rendering is handled by the caller (cosmic-text / font engine).
    }

    pub fn measure(&self) -> (f32, f32) { (self.width, self.height) }
}
