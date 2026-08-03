//! ToolStrip widget — toolbar with button-like items and separators.

use super::layout::{
    KeyEvent, LayoutRect, MouseButton as LayoutMouseButton, MouseEvent, MouseEventKind,
    PanelWidget, RenderContext, WidgetEvent, WidgetId };
use super::{WidgetColors, rounded_rect_path};
use tiny_skia::*;

#[derive(Clone, Debug)]
pub enum ToolStripItem {
    Button(String),
    Separator }

pub struct ToolStrip {
    pub items: Vec<ToolStripItem>,
    pub hover_index: Option<usize>,
    pub pressed_index: Option<usize>,
    pub item_size: f32,
    pub separator_width: f32,
    pub width: f32,
    pub height: f32,
    pub colors: WidgetColors,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent> }

impl ToolStrip {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            hover_index: None,
            pressed_index: None,
            item_size: 24.0,
            separator_width: 6.0,
            width: 400.0,
            height: 28.0,
            colors: WidgetColors {
                background: (245, 246, 247, 255),
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

    /// X position of each item (returns (x, width) pairs).
    pub fn item_positions(&self) -> Vec<(f32, f32)> {
        let mut positions = Vec::new();
        let mut cx = 2.0;
        for item in &self.items {
            match item {
                ToolStripItem::Button(_) => {
                    positions.push((cx, self.item_size));
                    cx += self.item_size + 1.0;
                }
                ToolStripItem::Separator => {
                    positions.push((cx, self.separator_width));
                    cx += self.separator_width;
                }
            }
        }
        positions
    }

    /// Paint — toolbar background, button items with hover/press, separators.
    /// Button label text/icons drawn by caller.
    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let ts = Transform::from_scale(scale, scale);
        let mut paint = Paint::default();
        paint.anti_alias = true;
        let mut stroke = Stroke::default();
        stroke.width = 1.0;

        // Toolbar background
        let (r, g, b, a) = self.colors.background;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(x, y, self.width, self.height) {
            pixmap.fill_rect(rect, &paint, ts, None);
        }

        // Bottom border
        paint.set_color_rgba8(210, 210, 210, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(x, y + self.height);
        pb.line_to(x + self.width, y + self.height);
        if let Some(path) = pb.finish() {
            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
        }

        let positions = self.item_positions();
        let btn_h = self.height - 4.0;
        let btn_y = y + 2.0;

        for (i, item) in self.items.iter().enumerate() {
            if i >= positions.len() {
                break;
            }
            let (ix, iw) = positions[i];
            let item_x = x + ix;

            match item {
                ToolStripItem::Button(_) => {
                    let is_pressed = self.pressed_index == Some(i);
                    let is_hovered = self.hover_index == Some(i);

                    if is_pressed {
                        // Pressed: darker background
                        paint.set_color_rgba8(200, 200, 200, 255);
                        if let Some(path) = rounded_rect_path(item_x, btn_y, iw, btn_h, 2.0) {
                            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
                        }
                        paint.set_color_rgba8(160, 160, 160, 255);
                        if let Some(path) = rounded_rect_path(item_x, btn_y, iw, btn_h, 2.0) {
                            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                        }
                    } else if is_hovered {
                        // Hover: light highlight with border
                        paint.set_color_rgba8(225, 230, 235, 255);
                        if let Some(path) = rounded_rect_path(item_x, btn_y, iw, btn_h, 2.0) {
                            pixmap.fill_path(&path, &paint, FillRule::Winding, ts, None);
                        }
                        paint.set_color_rgba8(180, 185, 190, 255);
                        if let Some(path) = rounded_rect_path(item_x, btn_y, iw, btn_h, 2.0) {
                            pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                        }
                    }
                }
                ToolStripItem::Separator => {
                    // Vertical separator line
                    paint.set_color_rgba8(190, 190, 190, 255);
                    let sep_x = item_x + iw / 2.0;
                    let mut pb = PathBuilder::new();
                    pb.move_to(sep_x, btn_y + 2.0);
                    pb.line_to(sep_x, btn_y + btn_h - 2.0);
                    if let Some(path) = pb.finish() {
                        pixmap.stroke_path(&path, &paint, &stroke, ts, None);
                    }
                }
            }
        }
    }

    pub fn measure(&self) -> (f32, f32) {
        (self.width, self.height)
    }

    /// Hit test — returns item index of the button at position.
    pub fn hit_test(&self, mx: f32, my: f32) -> Option<usize> {
        if my < 0.0 || my > self.height {
            return None;
        }
        let positions = self.item_positions();
        for (i, (ix, iw)) in positions.iter().enumerate() {
            if matches!(self.items.get(i), Some(ToolStripItem::Button(_))) {
                if mx >= *ix && mx < ix + iw {
                    return Some(i);
                }
            }
        }
        None
    }

    /// Update hover on mouse move.
    pub fn mouse_move(&mut self, mx: f32, my: f32) {
        self.hover_index = self.hit_test(mx, my);
    }
}

impl PanelWidget for ToolStrip {
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
        // Draw button labels
        let (fr, fg, fb, _) = self.colors.foreground;
        let col = cosmic_text::Color::rgba(fr, fg, fb, 255);
        let positions = self.item_positions();
        for (i, item) in self.items.iter().enumerate() {
            if i >= positions.len() {
                break;
            }
            if let ToolStripItem::Button(label) = item {
                let (ix, _iw) = positions[i];
                super::ide_text::draw_text(
                    ctx.pixmap,
                    ctx.font_system,
                    ctx.swash_cache,
                    label,
                    r.x + ix + 4.0,
                    r.y + 6.0,
                    11.0,
                    col,
                    ctx.scale,
                );
            }
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        if !r.contains(event.x, event.y) {
            self.hover_index = None;
            self.pressed_index = None;
            return false;
        }
        let lx = event.x - r.x;
        let ly = event.y - r.y;
        match event.kind {
            MouseEventKind::Move => {
                self.mouse_move(lx, ly);
            }
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                if let Some(idx) = self.hit_test(lx, ly) {
                    self.pressed_index = Some(idx);
                    self.pending_events
                        .push(WidgetEvent::ToolStripItemClicked(self.name.clone(), idx));
                    return true;
                }
            }
            MouseEventKind::Release(LayoutMouseButton::Left) => {
                self.pressed_index = None;
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
}
