use tiny_skia::{Paint, Pixmap, Transform, PathBuilder, Rect};
use cosmic_text::{FontSystem, SwashCache, Color as CosmicColor};

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
            col_w: 0.0,  // 0.0 triggers actual_col_w auto-calculation
            row_h: 25.0,
        }
    }

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

    pub fn render(&self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, x: f32, y: f32, bg: (u8, u8, u8, u8), border: (u8, u8, u8, u8), selection: (u8, u8, u8, u8), hover: (u8, u8, u8, u8), active_text: CosmicColor, inactive_text: CosmicColor) {
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

    pub fn handle_mouse(&mut self, mx: f32, my: f32, x: f32, y: f32, is_click: bool) -> DropdownEvent {
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
