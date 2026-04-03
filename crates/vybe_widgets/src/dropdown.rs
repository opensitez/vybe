use tiny_skia::{Paint, Pixmap, Transform, PathBuilder, Rect};
use cosmic_text::{FontSystem, SwashCache, Color as CosmicColor};
use super::layout::{LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton, KeyEvent, RenderContext, PanelWidget, WidgetEvent, WidgetId, WidgetCommand, CommandValue};

pub enum DropdownEvent {
    Selected(usize),
    Closed,
    None,
}

pub struct Dropdown {
    pub items: Vec<String>,
    pub selected_idx: usize,
    pub hover_idx: Option<usize>,
    pub scale: f32,
    pub num_cols: usize,
    pub col_w: f32,
    pub row_h: f32,
    pub id: WidgetId,
    pub name: String,
    rect: LayoutRect,
    pending_events: Vec<WidgetEvent>,
}

impl Dropdown {
    pub fn new(items: Vec<String>, selected_idx: usize, scale: f32, num_cols: Option<usize>) -> Self {
        let actual_cols = num_cols.unwrap_or_else(|| {
            if items.len() <= 15 { 1 }
            else { (items.len() as f32 / 15.0).ceil() as usize }
        }).max(1);
        
        Self {
            items,
            selected_idx,
            hover_idx: None,
            scale,
            num_cols: actual_cols,
            col_w: 0.0,
            row_h: 25.0,
            id: WidgetId::next(),
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    fn actual_col_w(&self) -> f32 {
        if self.col_w == 0.0 {
            let max_chars = self.items.iter().map(|s| s.len()).max().unwrap_or(10);
            (max_chars as f32 * 9.0 + 30.0).max(120.0)
        } else {
            self.col_w
        }
    }

    pub fn get_size(&self) -> (f32, f32) {
        let col_w = self.actual_col_w();
        let items_per_col = (self.items.len() as f32 / self.num_cols as f32).ceil() as usize;
        let w = self.num_cols as f32 * col_w + 10.0;
        let h = items_per_col as f32 * self.row_h + 10.0;
        (w, h)
    }

    pub fn render_list(&self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, x: f32, y: f32, bg: (u8, u8, u8, u8), border: (u8, u8, u8, u8), selection: (u8, u8, u8, u8), hover: (u8, u8, u8, u8), active_text: CosmicColor, inactive_text: CosmicColor) {
        let (w, h) = self.get_size();
        let scale = self.scale;
        
        let rw = w * scale;
        let rh = h * scale;
        let rx = x * scale;
        let ry = y * scale;

        // 1. Background
        let mut mp = Paint::default(); mp.set_color_rgba8(bg.0, bg.1, bg.2, bg.3);
        pix.fill_rect(Rect::from_xywh(rx, ry, rw, rh).unwrap(), &mp, Transform::identity(), None);

        // 2. Border
        let mut bp = Paint::default(); bp.set_color_rgba8(border.0, border.1, border.2, border.3);
        let mut pb = PathBuilder::new(); pb.push_rect(Rect::from_xywh(rx, ry, rw, rh).unwrap());
        if let Some(path) = pb.finish() {
            pix.stroke_path(&path, &bp, &tiny_skia::Stroke { width: 1.0 * scale, ..Default::default() }, Transform::identity(), None);
        }

        let col_w = self.actual_col_w();
        let items_per_col = (self.items.len() as f32 / self.num_cols as f32).ceil() as usize;
        for (i, item) in self.items.iter().enumerate() {
            let col = i / items_per_col;
            let row = i % items_per_col;
            
            let ix = rx + (5.0 + col as f32 * col_w) * scale;
            let iy = ry + (5.0 + row as f32 * self.row_h) * scale;
            let iw = col_w * scale;
            let ih = self.row_h * scale;

            // Hover / Selected background
            if Some(i) == self.hover_idx {
                let mut hp = Paint::default(); hp.set_color_rgba8(hover.0, hover.1, hover.2, hover.3);
                pix.fill_rect(Rect::from_xywh(ix, iy, iw, ih).unwrap(), &hp, Transform::identity(), None);
            } else if i == self.selected_idx {
                let mut sp = Paint::default(); sp.set_color_rgba8(selection.0, selection.1, selection.2, selection.3);
                pix.fill_rect(Rect::from_xywh(ix, iy, iw, ih).unwrap(), &sp, Transform::identity(), None);
            }

            let is_active = i == self.selected_idx;
            let text_col = if is_active { active_text } else { inactive_text };
            
            // Center text vertically: (ih - line_height)/2.0
            let it_y = iy + (ih - 18.0 * scale) / 2.0;
            crate::tree_view::TreeView::draw_text_static_internal(pix, fs, sc, item, ix + 5.0 * scale, it_y, text_col, scale);
        }
    }

    pub fn handle_mouse_at(&mut self, mx: f32, my: f32, x: f32, y: f32, is_click: bool) -> DropdownEvent {
        let (w, h) = self.get_size();
        let col_w = self.actual_col_w();
        if mx >= x && mx <= x + w && my >= y && my <= y + h {
            let items_per_col = (self.items.len() as f32 / self.num_cols as f32).ceil() as usize;
            let col_idx = ((mx - x - 5.0) / col_w) as usize;
            let row_idx = ((my - y - 5.0) / self.row_h) as usize;
            let idx = col_idx * items_per_col + row_idx;

            if idx < self.items.len() {
                self.hover_idx = Some(idx);
                if is_click {
                    return DropdownEvent::Selected(idx);
                }
            } else {
                self.hover_idx = None;
            }
            DropdownEvent::None
        } else {
            self.hover_idx = None;
            if is_click { DropdownEvent::Closed } else { DropdownEvent::None }
        }
    }
}

impl PanelWidget for Dropdown {
    fn name(&self) -> &str { &self.name }
    fn widget_id(&self) -> WidgetId { self.id }
    fn set_rect(&mut self, rect: LayoutRect) { self.rect = rect; }
    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }
        let bg = (50, 50, 55, 255);
        let border = (80, 80, 85, 255);
        let selection = (0, 120, 215, 255);
        let hover = (65, 65, 70, 255);
        let active_text = CosmicColor::rgba(255, 255, 255, 255);
        let inactive_text = CosmicColor::rgba(200, 200, 200, 255);
        self.render_list(ctx.pixmap, ctx.font_system, ctx.swash_cache, r.x, r.y, bg, border, selection, hover, active_text, inactive_text);
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        let r = self.rect;
        let is_click = matches!(event.kind, MouseEventKind::Press(LayoutMouseButton::Left));
        match self.handle_mouse_at(event.x, event.y, r.x, r.y, is_click) {
            DropdownEvent::Selected(idx) => {
                self.pending_events.push(WidgetEvent::DropdownSelected(self.name.clone(), idx));
                true
            }
            DropdownEvent::Closed => true,
            DropdownEvent::None => false,
        }
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }
    fn handle_command(&mut self, cmd: &WidgetCommand) -> CommandValue {
        match cmd {
            WidgetCommand::SetSelectedIndex(i) => { if *i < self.items.len() { self.selected_idx = *i; } CommandValue::None }
            WidgetCommand::GetValue => CommandValue::Index(self.selected_idx),
            WidgetCommand::AddItem(s) => { self.items.push(s.clone()); CommandValue::None }
            WidgetCommand::RemoveItem(i) => { if *i < self.items.len() { self.items.remove(*i); } CommandValue::None }
            WidgetCommand::ClearItems => { self.items.clear(); self.selected_idx = 0; CommandValue::None }
            WidgetCommand::GetText => {
                let t = self.items.get(self.selected_idx).cloned().unwrap_or_default();
                CommandValue::Text(t)
            }
            _ => CommandValue::None,
        }
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> { std::mem::take(&mut self.pending_events) }
}
