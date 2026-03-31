//! Panel widget — container with optional border.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum BorderStyle {
    None,
    FixedSingle,
    Fixed3D,
}

pub struct Panel {
    pub border_style: BorderStyle,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
}

impl Panel {
    pub fn new() -> Self {
        Self {
            border_style: BorderStyle::None,
            width: 200.0,
            height: 150.0,
            colors: WidgetColors {
                background: (240, 240, 240, 255),
                ..WidgetColors::default()
            },
        }
    }

    /// Paint the panel — light background with optional border.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        match self.border_style {
            BorderStyle::None => {}
            BorderStyle::FixedSingle => {
                paint.set_color_rgba8(160, 160, 160, 255);
                if let Some(path) = rounded_rect_path(x, y, self.width, self.height, 0.0) {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
            }
            BorderStyle::Fixed3D => {
                // 3D sunken border
                paint.set_color_rgba8(130, 135, 144, 255);
                let mut pb = PathBuilder::new();
                pb.move_to(x, y + self.height);
                pb.line_to(x, y);
                pb.line_to(x + self.width, y);
                if let Some(path) = pb.finish() {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
                paint.set_color_rgba8(255, 255, 255, 255);
                let mut pb = PathBuilder::new();
                pb.move_to(x + self.width, y);
                pb.line_to(x + self.width, y + self.height);
                pb.line_to(x, y + self.height);
                if let Some(path) = pb.finish() {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }
}
