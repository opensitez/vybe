use tiny_skia::{Pixmap, Paint, PathBuilder, Transform, Stroke};
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent, WidgetId};

pub struct Toolbox {
    pub items: Vec<String>,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl Toolbox {
    pub fn new(items: Vec<&str>) -> Self {
        Self { items: items.into_iter().map(|s| s.to_string()).collect(), id: WidgetId::next(), name: String::new(), rect: LayoutRect::zero(), pending_events: Vec::new() }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    pub fn paint(&self, pixmap: &mut Pixmap, x: f32, y: f32, scale: f32) {
        let mut p = Paint::default(); p.set_color_rgba8(40, 40, 40, 220);
        let w = 160.0 * scale; let h = (self.items.len() as f32 * 28.0 + 16.0) * scale;
        let mut pb = PathBuilder::new(); pb.push_rect(tiny_skia::Rect::from_xywh(x, y, w, h).unwrap());
        if let Some(path) = pb.finish() {
            pixmap.fill_path(&path, &p, tiny_skia::FillRule::Winding, Transform::identity(), None);
        }
        // Items
        let mut iy = y + 8.0 * scale;
        for _item in &self.items {
            let mut b = Paint::default(); b.set_color_rgba8(220, 220, 220, 255);
            // Draw label (simple rectangle placeholder for now)
            let mut pb2 = PathBuilder::new(); pb2.push_rect(tiny_skia::Rect::from_xywh(x + 8.0*scale, iy, w - 16.0*scale, 22.0*scale).unwrap());
            if let Some(pth) = pb2.finish() { pixmap.stroke_path(&pth, &b, &Stroke::default(), Transform::identity(), None); }
            iy += 28.0 * scale;
        }
    }

    /// Hit-test the toolbox items. Returns Some(index) if the point (cx,cy)
    /// falls within an item rect, otherwise None.
    pub fn hit_test(&self, cx: f32, cy: f32, x: f32, y: f32, scale: f32) -> Option<usize> {
        let w = 160.0 * scale;
        let mut iy = y + 8.0 * scale;
        for (i, _item) in self.items.iter().enumerate() {
            let rx = x + 8.0 * scale;
            let ry = iy;
            let rw = w - 16.0 * scale;
            let rh = 22.0 * scale;
            if cx >= rx && cx < rx + rw && cy >= ry && cy < ry + rh {
                return Some(i);
            }
            iy += 28.0 * scale;
        }
        None
    }
}

impl PanelWidget for Toolbox {
    fn name(&self) -> &str { &self.name }
    fn widget_id(&self) -> WidgetId { self.id }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; }
    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        self.paint(ctx.pixmap, r.x, r.y, ctx.scale);
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        if !r.contains(event.x, event.y) { return false; }
        if let MouseEventKind::Press(LayoutMouseButton::Left) = event.kind {
            if let Some(idx) = self.hit_test(event.x, event.y, r.x, r.y, 1.0) {
                self.pending_events.push(WidgetEvent::Action(format!("toolbox:{}:{}", self.name, idx)));
                return true;
            }
        }
        false
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
    fn drain_events(&mut self) -> Vec<WidgetEvent> { std::mem::take(&mut self.pending_events) }
}
