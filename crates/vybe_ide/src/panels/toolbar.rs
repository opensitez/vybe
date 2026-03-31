//! Toolbar panel — matches legacy editor: Run, Stop, View Object/Code, New, Remove.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform};

use crate::layout::Rect;
use crate::text::draw_text;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ToolbarAction {
    Run,
    Stop,
    ViewDesigner,
    ViewCode,
    AddForm,
    AddCode,
}

struct TButton {
    label: &'static str,
    action: Option<ToolbarAction>,
    is_sep: bool,
}

const BUTTONS: &[TButton] = &[
    TButton { label: "\u{25B6} Start", action: Some(ToolbarAction::Run),          is_sep: false },
    TButton { label: "\u{25A0} End",   action: Some(ToolbarAction::Stop),         is_sep: false },
    TButton { label: "",               action: None,                               is_sep: true },
    TButton { label: "Designer",       action: Some(ToolbarAction::ViewDesigner), is_sep: false },
    TButton { label: "Code",           action: Some(ToolbarAction::ViewCode),     is_sep: false },
    TButton { label: "",               action: None,                               is_sep: true },
    TButton { label: "+ Form",         action: Some(ToolbarAction::AddForm),      is_sep: false },
    TButton { label: "+ Code",         action: Some(ToolbarAction::AddCode),      is_sep: false },
];

pub struct Toolbar;

impl Toolbar {
    pub fn render(pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: Rect, scale: f32) {
        let s = scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(240, 240, 240, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        // Bottom border
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + rect.h - 1.0, rect.w, 1.0, s);

        let text_color = CosmicColor::rgba(50, 50, 50, 255);
        let mut x = rect.x + 8.0;
        let btn_h = 26.0;
        let btn_y = rect.y + (rect.h - btn_h) / 2.0;

        for btn in BUTTONS {
            if btn.is_sep {
                paint.set_color_rgba8(200, 200, 200, 255);
                fill(pix, &paint, x, btn_y + 3.0, 1.0, btn_h - 6.0, s);
                x += 10.0;
                continue;
            }

            let btn_w = btn.label.len() as f32 * 8.0 + 20.0;

            // Button bg
            paint.set_color_rgba8(255, 255, 255, 255);
            fill(pix, &paint, x, btn_y, btn_w, btn_h, s);

            // Button border
            paint.set_color_rgba8(204, 204, 204, 255);
            fill(pix, &paint, x, btn_y, btn_w, 1.0, s);
            fill(pix, &paint, x, btn_y + btn_h - 1.0, btn_w, 1.0, s);
            fill(pix, &paint, x, btn_y, 1.0, btn_h, s);
            fill(pix, &paint, x + btn_w - 1.0, btn_y, 1.0, btn_h, s);

            draw_text(pix, fs, sc, btn.label, x + 10.0, btn_y + 5.0, 13.0, text_color, s);
            x += btn_w + 4.0;
        }
    }

    pub fn handle_click(mx: f32, my: f32, rect: Rect) -> Option<ToolbarAction> {
        let btn_h = 26.0;
        let btn_y = rect.y + (rect.h - btn_h) / 2.0;
        if my < btn_y || my > btn_y + btn_h { return None; }

        let mut x = rect.x + 8.0;
        for btn in BUTTONS {
            if btn.is_sep { x += 10.0; continue; }
            let btn_w = btn.label.len() as f32 * 8.0 + 20.0;
            if mx >= x && mx < x + btn_w {
                return btn.action;
            }
            x += btn_w + 4.0;
        }
        None
    }
}

fn fill(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pix.fill_rect(r, paint, Transform::identity(), None);
    }
}
