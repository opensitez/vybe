//! ScrollBar widget — track with draggable thumb and arrow buttons.

use tiny_skia::*;
use super::WidgetColors;
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent};

pub struct ScrollBar {
    pub vertical: bool,
    pub pos: f32,
    pub content_size: f32,
    pub viewport_size: f32,
    pub dragging: bool,
    pub drag_offset: f32,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl ScrollBar {
    pub fn new(vertical: bool) -> Self {
        Self {
            vertical,
            pos: 0.0,
            content_size: 500.0,
            viewport_size: 200.0,
            dragging: false,
            drag_offset: 0.0,
            width: if vertical { 16.0 } else { 160.0 },
            height: if vertical { 160.0 } else { 16.0 },
            colors: WidgetColors::default(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    /// Arrow button size at each end.
    fn arrow_size(&self) -> f32 {
        if self.vertical { self.width } else { self.height }
    }

    /// Track length (excluding arrow buttons).
    fn track_length(&self) -> f32 {
        let total = if self.vertical { self.height } else { self.width };
        (total - self.arrow_size() * 2.0).max(0.0)
    }

    /// Thumb size proportional to viewport/content ratio.
    fn thumb_size(&self) -> f32 {
        if self.content_size <= 0.0 || self.viewport_size >= self.content_size {
            return self.track_length();
        }
        let ratio = self.viewport_size / self.content_size;
        (self.track_length() * ratio).max(20.0)
    }

    /// Thumb offset within track.
    fn thumb_offset(&self) -> f32 {
        let available = self.track_length() - self.thumb_size();
        if available <= 0.0 { return 0.0; }
        self.pos.clamp(0.0, 1.0) * available
    }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        let arrow = self.arrow_size();

        // Track background
        paint.set_color_rgba8(240, 240, 240, 255);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Track border
        paint.set_color_rgba8(200, 200, 200, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + self.width, y);
        pb.line_to(x + self.width, y + self.height);
        pb.line_to(x, y + self.height);
        pb.close();
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        if self.vertical {
            // Top arrow button
            paint.set_color_rgba8(228, 228, 228, 255);
            if let Some(rect) = Rect::from_xywh(x + 1.0, y + 1.0, self.width - 2.0, arrow - 1.0) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
            // Top arrow triangle (up)
            paint.set_color_rgba8(80, 80, 80, 255);
            let cx = x + self.width / 2.0;
            let cy = y + arrow / 2.0;
            let as_ = 3.0;
            let mut pb = PathBuilder::new();
            pb.move_to(cx, cy - as_);
            pb.line_to(cx + as_, cy + as_ * 0.6);
            pb.line_to(cx - as_, cy + as_ * 0.6);
            pb.close();
            if let Some(path) = pb.finish() {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }

            // Bottom arrow button
            paint.set_color_rgba8(228, 228, 228, 255);
            let bottom_y = y + self.height - arrow;
            if let Some(rect) = Rect::from_xywh(x + 1.0, bottom_y, self.width - 2.0, arrow - 1.0) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
            // Bottom arrow triangle (down)
            paint.set_color_rgba8(80, 80, 80, 255);
            let cy = bottom_y + arrow / 2.0;
            let mut pb = PathBuilder::new();
            pb.move_to(cx, cy + as_);
            pb.line_to(cx + as_, cy - as_ * 0.6);
            pb.line_to(cx - as_, cy - as_ * 0.6);
            pb.close();
            if let Some(path) = pb.finish() {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }

            // Dividers between arrows and track
            paint.set_color_rgba8(200, 200, 200, 255);
            let mut pb = PathBuilder::new();
            pb.move_to(x, y + arrow);
            pb.line_to(x + self.width, y + arrow);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
            let mut pb = PathBuilder::new();
            pb.move_to(x, bottom_y);
            pb.line_to(x + self.width, bottom_y);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Thumb
            let thumb_y = y + arrow + self.thumb_offset();
            let thumb_h = self.thumb_size();
            paint.set_color_rgba8(190, 190, 190, 255);
            if let Some(rect) = Rect::from_xywh(x + 2.0, thumb_y, self.width - 4.0, thumb_h) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
            // Thumb border
            paint.set_color_rgba8(160, 160, 160, 255);
            let mut pb = PathBuilder::new();
            pb.move_to(x + 2.0, thumb_y);
            pb.line_to(x + self.width - 2.0, thumb_y);
            pb.line_to(x + self.width - 2.0, thumb_y + thumb_h);
            pb.line_to(x + 2.0, thumb_y + thumb_h);
            pb.close();
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Grip dots on thumb (3 horizontal lines)
            paint.set_color_rgba8(140, 140, 140, 255);
            let grip_cx = x + self.width / 2.0;
            let grip_cy = thumb_y + thumb_h / 2.0;
            for i in -1..=1i32 {
                let gy = grip_cy + i as f32 * 3.0;
                let mut pb = PathBuilder::new();
                pb.move_to(grip_cx - 3.0, gy);
                pb.line_to(grip_cx + 3.0, gy);
                if let Some(path) = pb.finish() {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
            }
        } else {
            // Horizontal scrollbar
            // Left arrow button
            paint.set_color_rgba8(228, 228, 228, 255);
            if let Some(rect) = Rect::from_xywh(x + 1.0, y + 1.0, arrow - 1.0, self.height - 2.0) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
            paint.set_color_rgba8(80, 80, 80, 255);
            let cx = x + arrow / 2.0;
            let cy = y + self.height / 2.0;
            let as_ = 3.0;
            let mut pb = PathBuilder::new();
            pb.move_to(cx - as_, cy);
            pb.line_to(cx + as_ * 0.6, cy - as_);
            pb.line_to(cx + as_ * 0.6, cy + as_);
            pb.close();
            if let Some(path) = pb.finish() {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }

            // Right arrow button
            paint.set_color_rgba8(228, 228, 228, 255);
            let right_x = x + self.width - arrow;
            if let Some(rect) = Rect::from_xywh(right_x, y + 1.0, arrow - 1.0, self.height - 2.0) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
            paint.set_color_rgba8(80, 80, 80, 255);
            let cx = right_x + arrow / 2.0;
            let mut pb = PathBuilder::new();
            pb.move_to(cx + as_, cy);
            pb.line_to(cx - as_ * 0.6, cy - as_);
            pb.line_to(cx - as_ * 0.6, cy + as_);
            pb.close();
            if let Some(path) = pb.finish() {
                pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
            }

            // Dividers
            paint.set_color_rgba8(200, 200, 200, 255);
            let mut pb = PathBuilder::new();
            pb.move_to(x + arrow, y);
            pb.line_to(x + arrow, y + self.height);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }
            let mut pb = PathBuilder::new();
            pb.move_to(right_x, y);
            pb.line_to(right_x, y + self.height);
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Thumb
            let thumb_x = x + arrow + self.thumb_offset();
            let thumb_w = self.thumb_size();
            paint.set_color_rgba8(190, 190, 190, 255);
            if let Some(rect) = Rect::from_xywh(thumb_x, y + 2.0, thumb_w, self.height - 4.0) {
                pixmap.fill_rect(rect, &paint, ts, None);
            }
            paint.set_color_rgba8(160, 160, 160, 255);
            let mut pb = PathBuilder::new();
            pb.move_to(thumb_x, y + 2.0);
            pb.line_to(thumb_x + thumb_w, y + 2.0);
            pb.line_to(thumb_x + thumb_w, y + self.height - 2.0);
            pb.line_to(thumb_x, y + self.height - 2.0);
            pb.close();
            if let Some(path) = pb.finish() {
                pixmap.stroke_path(&path, &paint, &stroke, ts, None);
            }

            // Grip dots (3 vertical lines)
            paint.set_color_rgba8(140, 140, 140, 255);
            let grip_cx = thumb_x + thumb_w / 2.0;
            let grip_cy = y + self.height / 2.0;
            for i in -1..=1i32 {
                let gx = grip_cx + i as f32 * 3.0;
                let mut pb = PathBuilder::new();
                pb.move_to(gx, grip_cy - 3.0);
                pb.line_to(gx, grip_cy + 3.0);
                if let Some(path) = pb.finish() {
                    pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                }
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Handle mouse down — returns true if drag started on thumb.
    pub fn mouse_down(&mut self, mx: f32, my: f32) -> bool {
        let arrow = self.arrow_size();
        let thumb_offset = self.thumb_offset();
        let thumb_sz = self.thumb_size();

        if self.vertical {
            let thumb_y = arrow + thumb_offset;
            if mx >= 0.0 && mx <= self.width && my >= thumb_y && my <= thumb_y + thumb_sz {
                self.dragging = true;
                self.drag_offset = my - thumb_y;
                return true;
            }
        } else {
            let thumb_x = arrow + thumb_offset;
            if my >= 0.0 && my <= self.height && mx >= thumb_x && mx <= thumb_x + thumb_sz {
                self.dragging = true;
                self.drag_offset = mx - thumb_x;
                return true;
            }
        }
        false
    }

    /// Handle mouse move during drag.
    pub fn mouse_move(&mut self, mx: f32, my: f32) {
        if !self.dragging { return; }
        let arrow = self.arrow_size();
        let track = self.track_length();
        let thumb_sz = self.thumb_size();
        let available = track - thumb_sz;
        if available <= 0.0 { return; }

        let new_offset = if self.vertical {
            my - self.drag_offset - arrow
        } else {
            mx - self.drag_offset - arrow
        };
        self.pos = (new_offset / available).clamp(0.0, 1.0);
    }

    /// Handle mouse up — end drag.
    pub fn mouse_up(&mut self) {
        self.dragging = false;
    }

    /// Scroll offset in content pixels.
    pub fn scroll_offset(&self) -> f32 {
        if self.content_size <= self.viewport_size { return 0.0; }
        self.pos.clamp(0.0, 1.0) * (self.content_size - self.viewport_size)
    }
}

impl PanelWidget for ScrollBar {
    fn name(&self) -> &str { &self.name }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; self.width = rect.w; self.height = rect.h; }
    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        let lx = event.x - r.x;
        let ly = event.y - r.y;
        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                if r.contains(event.x, event.y) && self.mouse_down(lx, ly) {
                    self.pending_events.push(WidgetEvent::ScrollChanged(self.name.clone(), self.pos));
                    return true;
                }
            }
            MouseEventKind::Move => {
                if self.dragging {
                    self.mouse_move(lx, ly);
                    self.pending_events.push(WidgetEvent::ScrollChanged(self.name.clone(), self.pos));
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

    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
    fn drain_events(&mut self) -> Vec<WidgetEvent> { std::mem::take(&mut self.pending_events) }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        let r = self.rect;
        let lx = x - r.x;
        let ly = y - r.y;
        if r.contains(x, y) && self.vertical {
            winit::window::CursorIcon::NsResize
        } else if r.contains(x, y) {
            winit::window::CursorIcon::EwResize
        } else {
            winit::window::CursorIcon::Default
        }
    }
}
