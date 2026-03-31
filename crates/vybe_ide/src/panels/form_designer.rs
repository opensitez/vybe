//! Form designer panel — visual form editor with grid, snap, selection, resize.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform, Stroke, PathBuilder};
use uuid::Uuid;
use vybe_forms::{Form, Control, ControlType};

use crate::layout::Rect;
use crate::text::{draw_text, draw_text_with_font, measure_text_with_font};
use crate::panels::toolbox_panel::ControlTool;

fn next_ctrl_name(ct: &ControlType, form: &Form) -> String {
    // Find next available number for this type
    let prefix = format!("{:?}", ct);
    let mut max = 0u32;
    for ctrl in &form.controls {
        if ctrl.name.starts_with(&prefix) {
            if let Ok(n) = ctrl.name[prefix.len()..].parse::<u32>() {
                max = max.max(n);
            }
        }
    }
    format!("{}{}", prefix, max + 1)
}

/// Snap a value to a 10px grid.
fn snap(v: i32) -> i32 {
    (v / 10) * 10
}

const TITLE_H: f32 = 30.0;
const FORM_PADDING: f32 = 20.0;
const GRID_SIZE: f32 = 20.0;
const HANDLE_SZ: f32 = 6.0;

/// Which resize handle is being dragged.
#[derive(Clone, Copy, PartialEq)]
pub enum ResizeHandle {
    TopLeft, Top, TopRight,
    Left, Right,
    BottomLeft, Bottom, BottomRight,
}

pub struct FormDesigner {
    pub selected_controls: Vec<Uuid>,
    pub drag_start: Option<(f32, f32)>,
    pub drag_offset: Option<(f32, f32)>,
    pub dragging: bool,
    pub resize_handle: Option<ResizeHandle>,
    pub resize_initial: Option<(i32, i32, i32, i32)>, // x, y, w, h
    pub lasso_start: Option<(f32, f32)>,
    pub lasso_current: Option<(f32, f32)>,
    pub scroll_x: f32,
    pub scroll_y: f32,
}

impl FormDesigner {
    pub fn new() -> Self {
        Self {
            selected_controls: Vec::new(),
            drag_start: None,
            drag_offset: None,
            dragging: false,
            resize_handle: None,
            resize_initial: None,
            lasso_start: None,
            lasso_current: None,
            scroll_x: 0.0,
            scroll_y: 0.0,
        }
    }

    fn form_client_origin(&self, rect: Rect) -> (f32, f32) {
        (
            rect.x + FORM_PADDING - self.scroll_x,
            rect.y + FORM_PADDING - self.scroll_y + TITLE_H,
        )
    }

    pub fn render(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        rect: Rect, scale: f32, form: &Form,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        // Workspace background
        paint.set_color_rgba8(45, 45, 48, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        let form_x = rect.x + FORM_PADDING - self.scroll_x;
        let form_y = rect.y + FORM_PADDING - self.scroll_y;
        let form_w = form.width as f32;
        let form_h = form.height as f32;

        // Shadow
        paint.set_color_rgba8(0, 0, 0, 40);
        fill(pix, &paint, form_x + 3.0, form_y + 3.0, form_w, form_h, s);

        // Title bar (blue gradient-like)
        paint.set_color_rgba8(0, 120, 212, 255);
        fill(pix, &paint, form_x, form_y, form_w, TITLE_H, s);

        let title_color = CosmicColor::rgba(255, 255, 255, 255);
        draw_text(pix, fs, sc, &form.text, form_x + 10.0, form_y + 7.0, 13.0, title_color, s);

        // Window buttons
        let btn_y = form_y + 8.0;
        let btn_sz = 12.0;
        paint.set_color_rgba8(232, 77, 60, 255);
        fill(pix, &paint, form_x + form_w - 20.0, btn_y, btn_sz, btn_sz, s);
        paint.set_color_rgba8(241, 196, 15, 255);
        fill(pix, &paint, form_x + form_w - 38.0, btn_y, btn_sz, btn_sz, s);
        paint.set_color_rgba8(39, 174, 96, 255);
        fill(pix, &paint, form_x + form_w - 56.0, btn_y, btn_sz, btn_sz, s);

        // Client area
        paint.set_color_rgba8(240, 240, 240, 255);
        let client_y = form_y + TITLE_H;
        let client_h = form_h - TITLE_H;
        fill(pix, &paint, form_x, client_y, form_w, client_h, s);

        // Dot grid
        paint.set_color_rgba8(0, 0, 0, 30);
        let dot_sz = 1.5;
        let mut gx = 0.0;
        while gx < form_w {
            let mut gy = 0.0;
            while gy < client_h {
                fill(pix, &paint, form_x + gx, client_y + gy, dot_sz, dot_sz, s);
                gy += GRID_SIZE;
            }
            gx += GRID_SIZE;
        }

        // Form border
        paint.set_color_rgba8(100, 100, 100, 255);
        stroke_rect(pix, &paint, form_x, form_y, form_w, form_h, s);

        // Controls — render recursively (top-level first, then children inside containers)
        let (cx0, cy0) = self.form_client_origin(rect);
        self.render_controls_recursive(pix, fs, sc, form, None, cx0, cy0, s, 0);

        // Lasso
        if let (Some(start), Some(current)) = (self.lasso_start, self.lasso_current) {
            let lx = start.0.min(current.0);
            let ly = start.1.min(current.1);
            let lw = (start.0 - current.0).abs();
            let lh = (start.1 - current.1).abs();
            paint.set_color_rgba8(0, 102, 204, 30);
            fill(pix, &paint, lx, ly, lw, lh, s);
            paint.set_color_rgba8(0, 102, 204, 180);
            stroke_rect(pix, &paint, lx, ly, lw, lh, s);
        }

        // Component tray (non-visual controls)
        self.render_tray(pix, fs, sc, form, form_x, form_y + form_h, form_w, s);
    }

    // ── Component Tray ──

    const TRAY_H: f32 = 48.0;
    const TRAY_ITEM_W: f32 = 90.0;
    const TRAY_ITEM_H: f32 = 36.0;
    const TRAY_GAP: f32 = 6.0;

    fn non_visual_controls(form: &Form) -> Vec<&Control> {
        form.controls.iter().filter(|c| c.control_type.is_non_visual()).collect()
    }

    fn tray_icon(ct: &ControlType) -> &'static str {
        match ct {
            ControlType::BindingSourceComponent => "\u{1F517}",
            ControlType::BindingNavigator => "\u{1F9ED}",
            ControlType::DataSetComponent => "\u{1F5C4}",
            ControlType::DataTableComponent => "\u{1F4CB}",
            ControlType::DataAdapterComponent => "\u{1F50C}",
            ControlType::Timer => "\u{23F1}",
            ControlType::ImageList => "\u{1F5BC}",
            ControlType::ErrorProvider => "\u{26A0}",
            ControlType::ToolTip => "\u{1F4AC}",
            _ => "\u{2699}",
        }
    }

    fn render_tray(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        form: &Form, tray_x: f32, tray_top: f32, tray_w: f32, s: f32,
    ) {
        let non_visuals = Self::non_visual_controls(form);
        if non_visuals.is_empty() { return; }

        let mut paint = Paint::default();
        let tray_y = tray_top + 8.0;

        // Tray background
        paint.set_color_rgba8(245, 245, 245, 255);
        fill(pix, &paint, tray_x, tray_y, tray_w, Self::TRAY_H, s);

        // Tray top border
        paint.set_color_rgba8(200, 200, 200, 255);
        fill(pix, &paint, tray_x, tray_y, tray_w, 1.0, s);

        // Tray label
        let dim = CosmicColor::rgba(120, 120, 120, 255);
        draw_text(pix, fs, sc, "Components", tray_x + 4.0, tray_y + 2.0, 9.0, dim, s);

        // Items
        let text_color = CosmicColor::rgba(50, 50, 50, 255);
        let mut ix = tray_x + 4.0;
        let iy = tray_y + 12.0;

        for ctrl in &non_visuals {
            let is_sel = self.selected_controls.contains(&ctrl.id);

            // Background
            if is_sel {
                paint.set_color_rgba8(204, 228, 247, 255);
                fill(pix, &paint, ix, iy, Self::TRAY_ITEM_W, Self::TRAY_ITEM_H, s);
                paint.set_color_rgba8(0, 120, 212, 255);
                stroke_rect(pix, &paint, ix, iy, Self::TRAY_ITEM_W, Self::TRAY_ITEM_H, s);
            } else {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, ix, iy, Self::TRAY_ITEM_W, Self::TRAY_ITEM_H, s);
                paint.set_color_rgba8(210, 210, 210, 255);
                stroke_rect(pix, &paint, ix, iy, Self::TRAY_ITEM_W, Self::TRAY_ITEM_H, s);
            }

            // Icon
            let icon = Self::tray_icon(&ctrl.control_type);
            draw_text(pix, fs, sc, icon, ix + 4.0, iy + 4.0, 14.0, text_color, s);

            // Name (truncate if needed)
            let name = if ctrl.name.len() > 10 {
                format!("{}...", &ctrl.name[..8])
            } else {
                ctrl.name.clone()
            };
            draw_text(pix, fs, sc, &name, ix + 4.0, iy + 20.0, 9.0, text_color, s);

            ix += Self::TRAY_ITEM_W + Self::TRAY_GAP;
        }
    }

    /// Hit-test the component tray. Returns the control ID if clicked.
    fn hit_test_tray(&self, mx: f32, my: f32, rect: Rect, form: &Form) -> Option<Uuid> {
        let form_x = rect.x + FORM_PADDING - self.scroll_x;
        let form_y = rect.y + FORM_PADDING - self.scroll_y;
        let form_h = form.height as f32;
        let tray_y = form_y + form_h + 8.0 + 12.0; // +8 gap +12 label area

        let non_visuals = Self::non_visual_controls(form);
        if non_visuals.is_empty() { return None; }

        let mut ix = form_x + 4.0;
        for ctrl in &non_visuals {
            if mx >= ix && mx < ix + Self::TRAY_ITEM_W
                && my >= tray_y && my < tray_y + Self::TRAY_ITEM_H
            {
                return Some(ctrl.id);
            }
            ix += Self::TRAY_ITEM_W + Self::TRAY_GAP;
        }
        None
    }

    // ── Control Rendering ──

    fn is_container(ct: &ControlType) -> bool {
        matches!(ct,
            ControlType::Panel | ControlType::Frame | ControlType::PictureBox |
            ControlType::TabControl | ControlType::SplitContainer |
            ControlType::FlowLayoutPanel | ControlType::TableLayoutPanel
        )
    }

    /// Recursively render controls for a given parent.
    fn render_controls_recursive(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        form: &Form, parent_id: Option<Uuid>, offset_x: f32, offset_y: f32, s: f32, depth: usize,
    ) {
        if depth > 20 { return; }
        for ctrl in &form.controls {
            if ctrl.control_type.is_non_visual() { continue; }
            if ctrl.parent_id != parent_id { continue; }

            self.render_control(pix, fs, sc, ctrl, offset_x, offset_y, s);

            // If this is a container, recursively render children inside it
            if Self::is_container(&ctrl.control_type) {
                let child_x = offset_x + ctrl.bounds.x as f32;
                let child_y = offset_y + ctrl.bounds.y as f32;
                self.render_controls_recursive(
                    pix, fs, sc, form, Some(ctrl.id),
                    child_x, child_y, s, depth + 1,
                );
            }
        }
    }

    fn render_control(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        ctrl: &Control, offset_x: f32, offset_y: f32, s: f32,
    ) {
        let cx = offset_x + ctrl.bounds.x as f32;
        let cy = offset_y + ctrl.bounds.y as f32;
        let cw = ctrl.bounds.width as f32;
        let ch = ctrl.bounds.height as f32;
        let mut paint = Paint::default();
        let ctrl_text = ctrl.properties.get_string("Text").unwrap_or("").to_string();
        
        // Setup visual properties
        let font_prop = ctrl.properties.get_string("Font");
        
        let text_color = if let Some(hex) = ctrl.properties.get_string("ForeColor") {
            vybe_widgets::color_picker::PickedColor::from_hex(hex)
                .map(|c| CosmicColor::rgba(c.r, c.g, c.b, c.a))
                .unwrap_or(CosmicColor::rgba(30, 30, 30, 255))
        } else {
            CosmicColor::rgba(30, 30, 30, 255)
        };

        let back_color = if let Some(hex) = ctrl.properties.get_string("BackColor") {
            vybe_widgets::color_picker::PickedColor::from_hex(hex)
                .map(|c| (c.r, c.g, c.b, c.a))
        } else {
            None
        };

        match ctrl.control_type {
            ControlType::Button => {
                if let Some((r, g, b, a)) = back_color {
                    paint.set_color_rgba8(r, g, b, a);
                } else {
                    paint.set_color_rgba8(225, 225, 225, 255);
                }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(173, 173, 173, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                let tw = measure_text_with_font(fs, &ctrl_text, font_prop, 12.0, s);
                draw_text_with_font(pix, fs, sc, &ctrl_text, cx + (cw - tw) / 2.0, cy + (ch - 14.0) / 2.0, font_prop, 12.0, text_color, s);
            }
            ControlType::Label | ControlType::LinkLabel => {
                let color = if ctrl.control_type == ControlType::LinkLabel {
                    CosmicColor::rgba(0, 102, 204, 255)
                } else { text_color };
                if let Some((r, g, b, a)) = back_color {
                    paint.set_color_rgba8(r, g, b, a);
                    fill(pix, &paint, cx, cy, cw, ch, s);
                }
                draw_text_with_font(pix, fs, sc, &ctrl_text, cx + 2.0, cy + 2.0, font_prop, 12.0, color, s);
            }
            ControlType::TextBox | ControlType::MaskedTextBox => {
                if let Some((r, g, b, a)) = back_color {
                    paint.set_color_rgba8(r, g, b, a);
                } else {
                    paint.set_color_rgba8(255, 255, 255, 255);
                }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                let display = if ctrl_text.is_empty() { &ctrl.name } else { &ctrl_text };
                draw_text_with_font(pix, fs, sc, display, cx + 4.0, cy + 4.0, font_prop, 12.0, text_color, s);
            }
            ControlType::CheckBox => {
                let bsz = 14.0;
                let by = cy + (ch - bsz) / 2.0;
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx + 2.0, by, bsz, bsz, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx + 2.0, by, bsz, bsz, s);
                draw_text_with_font(pix, fs, sc, &ctrl_text, cx + 20.0, cy + 2.0, font_prop, 12.0, text_color, s);
            }
            ControlType::RadioButton => {
                paint.set_color_rgba8(122, 122, 122, 255);
                let cr = 7.0;
                if let Some(path) = vybe_widgets::circle_path((cx + 2.0 + cr) * s, (cy + ch / 2.0) * s, cr * s) {
                    let mut st = Stroke::default(); st.width = 1.0 * s;
                    pix.stroke_path(&path, &paint, &st, Transform::identity(), None);
                }
                draw_text_with_font(pix, fs, sc, &ctrl_text, cx + 20.0, cy + 2.0, font_prop, 12.0, text_color, s);
            }
            ControlType::ComboBox => {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(225, 225, 225, 255);
                fill(pix, &paint, cx + cw - 20.0, cy, 20.0, ch, s);
                // Arrow
                paint.set_color_rgba8(80, 80, 80, 255);
                let ax = cx + cw - 14.0;
                let ay = cy + ch / 2.0 - 2.0;
                let mut pb = PathBuilder::new();
                pb.move_to(ax * s, ay * s);
                pb.line_to((ax + 8.0) * s, ay * s);
                pb.line_to((ax + 4.0) * s, (ay + 5.0) * s);
                pb.close();
                if let Some(path) = pb.finish() {
                    pix.fill_path(&path, &paint, tiny_skia::FillRule::Winding, Transform::identity(), None);
                }
                draw_text_with_font(pix, fs, sc, &ctrl_text, cx + 4.0, cy + 4.0, font_prop, 12.0, text_color, s);
            }
            ControlType::ListBox => {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                // First item highlight
                paint.set_color_rgba8(0, 120, 212, 20);
                fill(pix, &paint, cx + 1.0, cy + 1.0, cw - 2.0, 18.0, s);
            }
            ControlType::Panel | ControlType::Frame => {
                paint.set_color_rgba8(236, 236, 236, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(160, 160, 160, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                if ctrl.control_type == ControlType::Frame {
                    paint.set_color_rgba8(236, 236, 236, 255);
                    fill(pix, &paint, cx + 6.0, cy - 2.0, ctrl_text.len() as f32 * 7.0 + 8.0, 14.0, s);
                    draw_text(pix, fs, sc, &ctrl_text, cx + 10.0, cy - 2.0, 12.0, text_color, s);
                }
            }
            ControlType::PictureBox => {
                paint.set_color_rgba8(210, 210, 210, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(160, 160, 160, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                // X cross for image placeholder
                let mut pb = PathBuilder::new();
                pb.move_to(cx * s, cy * s);
                pb.line_to((cx + cw) * s, (cy + ch) * s);
                pb.move_to((cx + cw) * s, cy * s);
                pb.line_to(cx * s, (cy + ch) * s);
                if let Some(path) = pb.finish() {
                    paint.set_color_rgba8(180, 180, 180, 255);
                    let mut st = Stroke::default(); st.width = 0.5 * s;
                    pix.stroke_path(&path, &paint, &st, Transform::identity(), None);
                }
            }
            ControlType::ProgressBar => {
                paint.set_color_rgba8(230, 230, 230, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(6, 176, 37, 255);
                fill(pix, &paint, cx + 1.0, cy + 1.0, (cw - 2.0) * 0.3, ch - 2.0, s);
                paint.set_color_rgba8(188, 188, 188, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::TrackBar => {
                let ty = cy + ch / 2.0 - 2.0;
                paint.set_color_rgba8(214, 214, 214, 255);
                fill(pix, &paint, cx + 8.0, ty, cw - 16.0, 4.0, s);
                paint.set_color_rgba8(0, 120, 212, 255);
                fill(pix, &paint, cx + cw * 0.3 - 4.0, cy + ch / 2.0 - 8.0, 8.0, 16.0, s);
            }
            ControlType::TabControl => {
                paint.set_color_rgba8(240, 240, 240, 255);
                fill(pix, &paint, cx, cy + 26.0, cw, ch - 26.0, s);
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, 80.0, 26.0, s);
                paint.set_color_rgba8(230, 230, 230, 255);
                fill(pix, &paint, cx + 82.0, cy, 80.0, 26.0, s);
                paint.set_color_rgba8(160, 160, 160, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                draw_text(pix, fs, sc, "Tab 1", cx + 12.0, cy + 5.0, 11.0, text_color, s);
                let grey = CosmicColor::rgba(100, 100, 100, 255);
                draw_text(pix, fs, sc, "Tab 2", cx + 94.0, cy + 5.0, 11.0, grey, s);
            }
            ControlType::DataGridView => {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                // Header row
                paint.set_color_rgba8(230, 230, 230, 255);
                fill(pix, &paint, cx, cy, cw, 22.0, s);
                let col_w = (cw / 3.0).floor();
                let grey = CosmicColor::rgba(80, 80, 80, 255);
                draw_text(pix, fs, sc, "Column1", cx + 4.0, cy + 3.0, 10.0, grey, s);
                draw_text(pix, fs, sc, "Column2", cx + col_w + 4.0, cy + 3.0, 10.0, grey, s);
                draw_text(pix, fs, sc, "Column3", cx + col_w * 2.0 + 4.0, cy + 3.0, 10.0, grey, s);
                // Grid lines
                paint.set_color_rgba8(220, 220, 220, 255);
                for i in 1..3 {
                    fill(pix, &paint, cx + col_w * i as f32, cy, 1.0, ch, s);
                }
                for i in 1..((ch / 20.0) as i32) {
                    fill(pix, &paint, cx, cy + i as f32 * 20.0, cw, 1.0, s);
                }
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::ListView => {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(230, 230, 230, 255);
                fill(pix, &paint, cx, cy, cw, 22.0, s);
                let grey = CosmicColor::rgba(80, 80, 80, 255);
                draw_text(pix, fs, sc, "Name", cx + 4.0, cy + 3.0, 10.0, grey, s);
                draw_text(pix, fs, sc, "Type", cx + cw * 0.4, cy + 3.0, 10.0, grey, s);
                draw_text(pix, fs, sc, "Size", cx + cw * 0.7, cy + 3.0, 10.0, grey, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::RichTextBox => {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::NumericUpDown => {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                draw_text(pix, fs, sc, "0", cx + 4.0, cy + 4.0, 12.0, text_color, s);
                // Spinner buttons
                paint.set_color_rgba8(225, 225, 225, 255);
                fill(pix, &paint, cx + cw - 18.0, cy, 18.0, ch, s);
                let grey = CosmicColor::rgba(80, 80, 80, 255);
                draw_text(pix, fs, sc, "\u{25B2}", cx + cw - 14.0, cy + 1.0, 8.0, grey, s);
                draw_text(pix, fs, sc, "\u{25BC}", cx + cw - 14.0, cy + ch / 2.0, 8.0, grey, s);
            }
            ControlType::DateTimePicker => {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                draw_text(pix, fs, sc, "01/01/2026", cx + 4.0, cy + 4.0, 12.0, text_color, s);
                paint.set_color_rgba8(225, 225, 225, 255);
                fill(pix, &paint, cx + cw - 20.0, cy, 20.0, ch, s);
            }
            ControlType::MenuStrip => {
                paint.set_color_rgba8(240, 240, 240, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                let grey = CosmicColor::rgba(50, 50, 50, 255);
                draw_text(pix, fs, sc, "File  Edit  View  Help", cx + 8.0, cy + 3.0, 11.0, grey, s);
            }
            ControlType::StatusStrip => {
                paint.set_color_rgba8(0, 120, 212, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                let white = CosmicColor::rgba(255, 255, 255, 255);
                draw_text(pix, fs, sc, "Ready", cx + 8.0, cy + 2.0, 11.0, white, s);
            }
            ControlType::SplitContainer => {
                paint.set_color_rgba8(240, 240, 240, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(180, 180, 180, 255);
                fill(pix, &paint, cx + cw / 2.0 - 2.0, cy, 4.0, ch, s);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::MonthCalendar => {
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(0, 120, 212, 255);
                fill(pix, &paint, cx, cy, cw, 24.0, s);
                let white = CosmicColor::rgba(255, 255, 255, 255);
                draw_text(pix, fs, sc, "March 2026", cx + 4.0, cy + 4.0, 12.0, white, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            _ => {
                if let Some((r, g, b, a)) = back_color {
                    paint.set_color_rgba8(r, g, b, a);
                } else {
                    paint.set_color_rgba8(230, 230, 230, 255);
                }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(160, 160, 160, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                if !ctrl_text.is_empty() {
                    draw_text_with_font(pix, fs, sc, &ctrl_text, cx + 4.0, cy + 4.0, font_prop, 12.0, text_color, s);
                } else {
                    let grey = CosmicColor::rgba(150, 150, 150, 255);
                    draw_text_with_font(pix, fs, sc, &ctrl.name, cx + 4.0, cy + 4.0, font_prop, 11.0, grey, s);
                }
            }
        }

        // Selection handles
        if self.selected_controls.contains(&ctrl.id) {
            self.render_selection_handles(pix, cx, cy, cw, ch, s);
        }
    }

    fn render_selection_handles(&self, pix: &mut Pixmap, x: f32, y: f32, w: f32, h: f32, s: f32) {
        let mut paint = Paint::default();

        // Dashed selection border
        paint.set_color_rgba8(0, 120, 212, 255);
        let mut stroke = Stroke::default();
        stroke.width = 2.0 * s;
        stroke.dash = tiny_skia::StrokeDash::new(vec![4.0 * s, 3.0 * s], 0.0);
        let mut pb = PathBuilder::new();
        if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
            pb.push_rect(r);
        }
        if let Some(path) = pb.finish() {
            pix.stroke_path(&path, &paint, &stroke, Transform::identity(), None);
        }

        // 8 handles (only for single selection)
        if self.selected_controls.len() == 1 {
            let half = HANDLE_SZ / 2.0;
            let handles = [
                (x - half, y - half), (x + w / 2.0 - half, y - half), (x + w - half, y - half),
                (x - half, y + h / 2.0 - half), (x + w - half, y + h / 2.0 - half),
                (x - half, y + h - half), (x + w / 2.0 - half, y + h - half), (x + w - half, y + h - half),
            ];
            for (hx, hy) in &handles {
                paint.set_color_rgba8(0, 120, 212, 255);
                fill(pix, &paint, *hx, *hy, HANDLE_SZ, HANDLE_SZ, s);
            }
        }
    }

    /// Place a new control from toolbox at canvas position.
    /// If placed inside a container, sets parent_id and adjusts coordinates to local.
    pub fn place_control(&mut self, mx: f32, my: f32, rect: Rect, form: &mut Form, tool: ControlTool) -> bool {
        let ct = match tool {
            ControlTool::Control(ct) => ct,
            ControlTool::Pointer => return false,
        };
        let (cx0, cy0) = self.form_client_origin(rect);

        // Find which container (if any) the click lands in
        let parent_id = Self::find_container_at(form, None, cx0, cy0, mx, my, 0);

        // Compute local position relative to parent
        let (parent_gx, parent_gy) = if let Some(pid) = parent_id {
            Self::compute_global_pos(form, pid, cx0, cy0)
        } else {
            (cx0, cy0)
        };

        let local_x = snap(((mx - parent_gx).max(0.0)) as i32);
        let local_y = snap(((my - parent_gy).max(0.0)) as i32);

        let name = next_ctrl_name(&ct, form);
        let mut ctrl = Control::new(ct.clone(), name, local_x, local_y);
        ctrl.parent_id = parent_id;
        let id = ctrl.id;
        form.controls.push(ctrl);
        self.selected_controls = vec![id];
        true
    }

    /// Find the deepest container control at (mx, my).
    fn find_container_at(
        form: &Form, parent_id: Option<Uuid>, offset_x: f32, offset_y: f32,
        mx: f32, my: f32, depth: usize,
    ) -> Option<Uuid> {
        if depth > 20 { return None; }
        let children: Vec<&Control> = form.controls.iter()
            .filter(|c| c.parent_id == parent_id && !c.control_type.is_non_visual())
            .collect();

        for ctrl in children.iter().rev() {
            let cx = offset_x + ctrl.bounds.x as f32;
            let cy = offset_y + ctrl.bounds.y as f32;
            let cw = ctrl.bounds.width as f32;
            let ch = ctrl.bounds.height as f32;

            if mx >= cx && mx < cx + cw && my >= cy && my < cy + ch {
                if Self::is_container(&ctrl.control_type) {
                    // Check deeper
                    if let Some(deeper) = Self::find_container_at(
                        form, Some(ctrl.id), cx, cy, mx, my, depth + 1,
                    ) {
                        return Some(deeper);
                    }
                    return Some(ctrl.id);
                }
            }
        }
        parent_id // return the current container (or None for form root)
    }

    /// Add a non-visual component directly (no canvas click needed).
    pub fn add_non_visual(&mut self, ct: ControlType, form: &mut Form) {
        let name = next_ctrl_name(&ct, form);
        let ctrl = Control::new(ct, name, 0, 0);
        let id = ctrl.id;
        form.controls.push(ctrl);
        self.selected_controls = vec![id];
    }

    /// Check if click is on a resize handle. Returns the handle if so.
    fn hit_test_handle(&self, mx: f32, my: f32, rect: Rect, form: &Form) -> Option<(Uuid, ResizeHandle)> {
        if self.selected_controls.len() != 1 { return None; }
        let id = self.selected_controls[0];
        let ctrl = form.controls.iter().find(|c| c.id == id)?;
        if ctrl.control_type.is_non_visual() { return None; }
        let (cx0, cy0) = self.form_client_origin(rect);
        let global = Self::compute_global_pos(form, id, cx0, cy0);
        let x = global.0;
        let y = global.1;
        let w = ctrl.bounds.width as f32;
        let h = ctrl.bounds.height as f32;
        let half = HANDLE_SZ / 2.0 + 2.0; // hit area slightly larger

        let handles = [
            (x - half, y - half, ResizeHandle::TopLeft),
            (x + w / 2.0 - half, y - half, ResizeHandle::Top),
            (x + w - half, y - half, ResizeHandle::TopRight),
            (x - half, y + h / 2.0 - half, ResizeHandle::Left),
            (x + w - half, y + h / 2.0 - half, ResizeHandle::Right),
            (x - half, y + h - half, ResizeHandle::BottomLeft),
            (x + w / 2.0 - half, y + h - half, ResizeHandle::Bottom),
            (x + w - half, y + h - half, ResizeHandle::BottomRight),
        ];
        let hs = HANDLE_SZ + 4.0;
        for (hx, hy, handle) in &handles {
            if mx >= *hx && mx < hx + hs && my >= *hy && my < hy + hs {
                return Some((id, *handle));
            }
        }
        None
    }

    /// Recursive hit-test. Returns the deepest control ID that contains (mx, my).
    fn hit_test_controls(
        form: &Form, parent_id: Option<Uuid>, offset_x: f32, offset_y: f32,
        mx: f32, my: f32, depth: usize,
    ) -> Option<Uuid> {
        if depth > 20 { return None; }

        // Check children in reverse order (top-most first), and for containers
        // check their children first (deepest wins).
        let children: Vec<&Control> = form.controls.iter()
            .filter(|c| c.parent_id == parent_id && !c.control_type.is_non_visual())
            .collect();

        for ctrl in children.iter().rev() {
            let cx = offset_x + ctrl.bounds.x as f32;
            let cy = offset_y + ctrl.bounds.y as f32;
            let cw = ctrl.bounds.width as f32;
            let ch = ctrl.bounds.height as f32;

            if mx >= cx && mx < cx + cw && my >= cy && my < cy + ch {
                // If container, check children first
                if Self::is_container(&ctrl.control_type) {
                    if let Some(child_hit) = Self::hit_test_controls(
                        form, Some(ctrl.id), cx, cy, mx, my, depth + 1,
                    ) {
                        return Some(child_hit);
                    }
                }
                return Some(ctrl.id);
            }
        }
        None
    }

    /// Compute the global (screen) position of a control by walking up the parent chain.
    fn compute_global_pos(form: &Form, ctrl_id: Uuid, form_x: f32, form_y: f32) -> (f32, f32) {
        let mut x = form_x;
        let mut y = form_y;
        let mut current_id = Some(ctrl_id);
        // Walk from the control up to root, collecting offsets
        let mut offsets = Vec::new();
        while let Some(cid) = current_id {
            if let Some(ctrl) = form.controls.iter().find(|c| c.id == cid) {
                offsets.push((ctrl.bounds.x as f32, ctrl.bounds.y as f32));
                current_id = ctrl.parent_id;
            } else {
                break;
            }
        }
        // Sum all offsets (root-most to self)
        for (dx, dy) in &offsets {
            x += dx;
            y += dy;
        }
        (x, y)
    }

    pub fn handle_mouse_down(&mut self, mx: f32, my: f32, rect: Rect, form: &Form, ctrl_held: bool) -> bool {
        if !rect.contains(mx, my) { return false; }

        // Check resize handles first
        if let Some((id, handle)) = self.hit_test_handle(mx, my, rect, form) {
            if let Some(ctrl) = form.controls.iter().find(|c| c.id == id) {
                self.resize_handle = Some(handle);
                self.resize_initial = Some((ctrl.bounds.x, ctrl.bounds.y, ctrl.bounds.width, ctrl.bounds.height));
                self.drag_start = Some((mx, my));
            }
            return true;
        }

        let (cx0, cy0) = self.form_client_origin(rect);

        // Recursive hit-test (deepest child wins)
        if let Some(hit_id) = Self::hit_test_controls(form, None, cx0, cy0, mx, my, 0) {
            let global = Self::compute_global_pos(form, hit_id, cx0, cy0);
            let cx = global.0;
            let cy = global.1;

            if ctrl_held {
                if let Some(pos) = self.selected_controls.iter().position(|&id| id == hit_id) {
                    self.selected_controls.remove(pos);
                } else {
                    self.selected_controls.push(hit_id);
                }
            } else if !self.selected_controls.contains(&hit_id) {
                self.selected_controls = vec![hit_id];
            }
            self.drag_start = Some((mx, my));
            self.drag_offset = Some((mx - cx, my - cy));
            self.dragging = false;
            return true;
        }

        // Check component tray
        if let Some(id) = self.hit_test_tray(mx, my, rect, form) {
            if ctrl_held {
                if let Some(pos) = self.selected_controls.iter().position(|&cid| cid == id) {
                    self.selected_controls.remove(pos);
                } else {
                    self.selected_controls.push(id);
                }
            } else {
                self.selected_controls = vec![id];
            }
            return true;
        }

        // Empty space — start lasso or deselect
        if !ctrl_held {
            self.selected_controls.clear();
        }
        self.lasso_start = Some((mx, my));
        self.lasso_current = Some((mx, my));
        true
    }

    pub fn handle_mouse_move(&mut self, mx: f32, my: f32, rect: Rect, form: &mut Form) {
        // Lasso
        if self.lasso_start.is_some() {
            self.lasso_current = Some((mx, my));
            return;
        }

        // Resize
        if let (Some(handle), Some(initial), Some(start)) = (self.resize_handle, self.resize_initial, self.drag_start) {
            let dx = (mx - start.0) as i32;
            let dy = (my - start.1) as i32;
            let (ix, iy, iw, ih) = initial;
            let id = self.selected_controls[0];
            if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == id) {
                let (mut nx, mut ny, mut nw, mut nh) = (ix, iy, iw, ih);
                match handle {
                    ResizeHandle::Right => { nw = (iw + dx).max(10); }
                    ResizeHandle::Bottom => { nh = (ih + dy).max(10); }
                    ResizeHandle::BottomRight => { nw = (iw + dx).max(10); nh = (ih + dy).max(10); }
                    ResizeHandle::Left => { nx = ix + dx; nw = (iw - dx).max(10); }
                    ResizeHandle::Top => { ny = iy + dy; nh = (ih - dy).max(10); }
                    ResizeHandle::TopLeft => { nx = ix + dx; ny = iy + dy; nw = (iw - dx).max(10); nh = (ih - dy).max(10); }
                    ResizeHandle::TopRight => { ny = iy + dy; nw = (iw + dx).max(10); nh = (ih - dy).max(10); }
                    ResizeHandle::BottomLeft => { nx = ix + dx; nw = (iw - dx).max(10); nh = (ih + dy).max(10); }
                }
                ctrl.bounds.x = snap(nx);
                ctrl.bounds.y = snap(ny);
                ctrl.bounds.width = snap(nw).max(10);
                ctrl.bounds.height = snap(nh).max(10);
            }
            return;
        }

        // Drag
        if let (Some(start), Some(offset)) = (self.drag_start, self.drag_offset) {
            let dx = (mx - start.0).abs();
            let dy = (my - start.1).abs();
            if !self.dragging && (dx > 5.0 || dy > 5.0) {
                self.dragging = true;
            }
            if self.dragging {
                let (cx0, cy0) = self.form_client_origin(rect);
                let sel = self.selected_controls.clone();

                // Pre-compute parent global positions for selected controls
                let parent_globals: Vec<(Uuid, f32, f32)> = sel.iter().filter_map(|&id| {
                    let ctrl = form.controls.iter().find(|c| c.id == id)?;
                    let (pgx, pgy) = if let Some(pid) = ctrl.parent_id {
                        Self::compute_global_pos(form, pid, cx0, cy0)
                    } else {
                        (cx0, cy0)
                    };
                    Some((id, pgx, pgy))
                }).collect();

                for (id, pgx, pgy) in &parent_globals {
                    if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == *id) {
                        ctrl.bounds.x = snap(((mx - offset.0) - pgx) as i32);
                        ctrl.bounds.y = snap(((my - offset.1) - pgy) as i32);
                    }
                }
            }
        }
    }

    pub fn handle_mouse_up(&mut self, rect: Rect, form: &Form) {
        // Finalize lasso
        if let (Some(start), Some(end)) = (self.lasso_start, self.lasso_current) {
            let (cx0, cy0) = self.form_client_origin(rect);
            let lx = start.0.min(end.0);
            let ly = start.1.min(end.1);
            let lw = (start.0 - end.0).abs();
            let lh = (start.1 - end.1).abs();

            if lw > 3.0 || lh > 3.0 {
                self.selected_controls.clear();
                for ctrl in &form.controls {
                    if ctrl.control_type.is_non_visual() { continue; }
                    let ccx = cx0 + ctrl.bounds.x as f32;
                    let ccy = cy0 + ctrl.bounds.y as f32;
                    let ccw = ctrl.bounds.width as f32;
                    let cch = ctrl.bounds.height as f32;
                    if ccx + ccw > lx && ccx < lx + lw && ccy + cch > ly && ccy < ly + lh {
                        self.selected_controls.push(ctrl.id);
                    }
                }
            }
        }

        self.lasso_start = None;
        self.lasso_current = None;
        self.drag_start = None;
        self.drag_offset = None;
        self.dragging = false;
        self.resize_handle = None;
        self.resize_initial = None;
    }

    pub fn selected_control_name<'a>(&self, form: &'a Form) -> Option<&'a str> {
        if self.selected_controls.len() == 1 {
            form.controls.iter()
                .find(|c| c.id == self.selected_controls[0])
                .map(|c| c.name.as_str())
        } else {
            None
        }
    }
}

fn fill(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pix.fill_rect(r, paint, Transform::identity(), None);
    }
}

fn stroke_rect(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    let mut pb = PathBuilder::new();
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pb.push_rect(r);
    }
    if let Some(path) = pb.finish() {
        let mut stroke = Stroke::default();
        stroke.width = 1.0 * s;
        pix.stroke_path(&path, paint, &stroke, Transform::identity(), None);
    }
}
