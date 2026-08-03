//! SplitContainer widget — two panels with a draggable divider.

use super::WidgetColors;
use super::layout::{
    KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent, MouseEventKind,
    PanelWidget, RenderContext, WidgetEvent, WidgetId };
use tiny_skia::*;

pub struct SplitContainer {
    pub horizontal: bool,
    pub split_position: f32, // 0.0..1.0
    pub dragging: bool,
    pub splitter_width: f32,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent> }

impl SplitContainer {
    pub fn new(horizontal: bool) -> Self {
        Self {
            horizontal,
            split_position: 0.5,
            dragging: false,
            splitter_width: 6.0,
            width: 300.0,
            height: 200.0,
            colors: WidgetColors {
                background: (240, 240, 240, 255),
                ..WidgetColors::default()
            },
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new() }
    }

    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Pixel position of the splitter center.
    fn splitter_pos(&self) -> f32 {
        if self.horizontal {
            self.split_position * self.width
        } else {
            self.split_position * self.height
        }
    }

    /// Paint the split container — two panel backgrounds with divider bar and grip dots.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let sw = self.splitter_width;
        let sp = self.splitter_pos();

        // Panel 1 background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if self.horizontal {
            if let Some(rect) = Rect::from_xywh(x, y, sp - sw / 2.0, self.height) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
        } else {
            if let Some(rect) = Rect::from_xywh(x, y, self.width, sp - sw / 2.0) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
        }

        // Panel 2 background
        if self.horizontal {
            let p2_x = x + sp + sw / 2.0;
            if let Some(rect) = Rect::from_xywh(p2_x, y, self.width - sp - sw / 2.0, self.height) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
        } else {
            let p2_y = y + sp + sw / 2.0;
            if let Some(rect) = Rect::from_xywh(x, p2_y, self.width, self.height - sp - sw / 2.0) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
        }

        // Splitter bar
        paint.set_color_rgba8(215, 215, 215, 255);
        if self.horizontal {
            let sx = x + sp - sw / 2.0;
            if let Some(rect) = Rect::from_xywh(sx, y, sw, self.height) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
            // Splitter border lines
            paint.set_color_rgba8(190, 190, 190, 255);
            let mut pb = PathBuilder::new();
            pb.move_to(sx, y);
            pb.line_to(sx, y + self.height);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
            let mut pb = PathBuilder::new();
            pb.move_to(sx + sw, y);
            pb.line_to(sx + sw, y + self.height);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Grip dots (vertical dots in center)
            paint.set_color_rgba8(140, 140, 140, 255);
            let gcx = sx + sw / 2.0;
            let gcy = y + self.height / 2.0;
            for i in -2..=2i32 {
                if let Some(path) = super::circle_path(gcx, gcy + i as f32 * 4.0, 1.0) {
                    pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
                }
            }
        } else {
            let sy = y + sp - sw / 2.0;
            if let Some(rect) = Rect::from_xywh(x, sy, self.width, sw) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
            paint.set_color_rgba8(190, 190, 190, 255);
            let mut pb = PathBuilder::new();
            pb.move_to(x, sy);
            pb.line_to(x + self.width, sy);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
            let mut pb = PathBuilder::new();
            pb.move_to(x, sy + sw);
            pb.line_to(x + self.width, sy + sw);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Grip dots (horizontal dots in center)
            paint.set_color_rgba8(140, 140, 140, 255);
            let gcx = x + self.width / 2.0;
            let gcy = sy + sw / 2.0;
            for i in -2..=2i32 {
                if let Some(path) = super::circle_path(gcx + i as f32 * 4.0, gcy, 1.0) {
                    pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
                }
            }
        }

        // Outer border
        paint.set_color_rgba8(180, 180, 180, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Test if mouse is on the splitter bar (for cursor change).
    pub fn hit_splitter(&self, mx: f32, my: f32) -> bool {
        let sp = self.splitter_pos();
        let sw = self.splitter_width;
        if self.horizontal {
            mx >= sp - sw / 2.0 && mx <= sp + sw / 2.0 && my >= 0.0 && my <= self.height
        } else {
            my >= sp - sw / 2.0 && my <= sp + sw / 2.0 && mx >= 0.0 && mx <= self.width
        }
    }

    /// Begin dragging the splitter.
    pub fn mouse_down(&mut self, mx: f32, my: f32) -> bool {
        if self.hit_splitter(mx, my) {
            self.dragging = true;
            return true;
        }
        false
    }

    /// Update split position during drag.
    pub fn mouse_move(&mut self, mx: f32, my: f32) {
        if !self.dragging {
            return;
        }
        if self.horizontal {
            self.split_position = (mx / self.width).clamp(0.1, 0.9);
        } else {
            self.split_position = (my / self.height).clamp(0.1, 0.9);
        }
    }

    /// End drag.
    pub fn mouse_up(&mut self) {
        self.dragging = false;
    }

    /// Get the two panel rects: (panel1_x, panel1_y, panel1_w, panel1_h, panel2_x, panel2_y, panel2_w, panel2_h).
    pub fn panel_rects(&self) -> ((f32, f32, f32, f32), (f32, f32, f32, f32)) {
        let sp = self.splitter_pos();
        let sw = self.splitter_width;
        if self.horizontal {
            (
                (0.0, 0.0, sp - sw / 2.0, self.height),
                (sp + sw / 2.0, 0.0, self.width - sp - sw / 2.0, self.height),
            )
        } else {
            (
                (0.0, 0.0, self.width, sp - sw / 2.0),
                (0.0, sp + sw / 2.0, self.width, self.height - sp - sw / 2.0),
            )
        }
    }
}

impl PanelWidget for SplitContainer {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.width = rect.w;
        self.height = rect.h;
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        let lx = event.x - r.x;
        let ly = event.y - r.y;
        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                if r.contains(event.x, event.y) && self.mouse_down(lx, ly) {
                    return true;
                }
            }
            MouseEventKind::Move => {
                if self.dragging {
                    self.mouse_move(lx, ly);
                    self.pending_events.push(WidgetEvent::SplitMoved(
                        self.name.clone(),
                        self.split_position,
                    ));
                    return true;
                }
            }
            MouseEventKind::Release(LayoutMouseButton::Left) => {
                if self.dragging {
                    self.mouse_up();
                    return true;
                }
            }
            _ => {}
        }
        false
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }
    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        let r = self.rect;
        let lx = x - r.x;
        let ly = y - r.y;
        if self.hit_splitter(lx, ly) || self.dragging {
            if self.horizontal {
                winit::window::CursorIcon::EwResize
            } else {
                winit::window::CursorIcon::NsResize
            }
        } else {
            winit::window::CursorIcon::Default
        }
    }
}
