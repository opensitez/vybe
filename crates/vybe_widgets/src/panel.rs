//! Panel widget — container with optional border.

use tiny_skia::*;
use super::{WidgetColors, rounded_rect_path};
use super::layout::{LayoutRect, MouseEvent, KeyEvent, RenderContext, PanelWidget, WidgetId, WidgetCommand, CommandValue};

/// Parse a color string like "#C8C8C8" or "Red" into (r, g, b, a).
fn parse_color_string(s: &str) -> Option<(u8, u8, u8, u8)> {
    let s = s.trim();
    if s.starts_with('#') && s.len() >= 7 {
        let r = u8::from_str_radix(&s[1..3], 16).ok()?;
        let g = u8::from_str_radix(&s[3..5], 16).ok()?;
        let b = u8::from_str_radix(&s[5..7], 16).ok()?;
        return Some((r, g, b, 255));
    }
    match s.to_lowercase().as_str() {
        "red" => Some((255, 0, 0, 255)),
        "green" => Some((0, 128, 0, 255)),
        "blue" => Some((0, 0, 255, 255)),
        "white" => Some((255, 255, 255, 255)),
        "black" => Some((0, 0, 0, 255)),
        "yellow" => Some((255, 255, 0, 255)),
        "gray" | "grey" => Some((128, 128, 128, 255)),
        "lightgray" | "lightgrey" => Some((211, 211, 211, 255)),
        "transparent" => Some((0, 0, 0, 0)),
        _ => None,
    }
}

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
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
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
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

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

impl PanelWidget for Panel {
    fn name(&self) -> &str { &self.name }
    fn widget_id(&self) -> WidgetId { self.id }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; self.width = rect.w; self.height = rect.h; }
    fn rect(&self) -> LayoutRect { self.rect }
    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
    }
    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool { false }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::Custom(key, val) => {
                match key.as_str() {
                    "SetBackColor" => {
                        if let CommandValue::Text(color_str) = val {
                            if let Some(rgba) = parse_color_string(color_str) {
                                self.colors.background = rgba;
                            }
                        }
                        CommandValue::None
                    }
                    "SetBorderStyle" => {
                        if let CommandValue::Text(s) = val {
                            match s.to_lowercase().as_str() {
                                "fixedsingle" | "1" => self.border_style = BorderStyle::FixedSingle,
                                "fixed3d" | "2" => self.border_style = BorderStyle::Fixed3D,
                                _ => self.border_style = BorderStyle::None,
                            }
                        }
                        CommandValue::None
                    }
                    _ => CommandValue::None,
                }
            }
            _ => CommandValue::None,
        }
    }
}
