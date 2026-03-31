//! Font Picker widget — standalone tiny-skia rendered font family + size picker.
//!
//! Renders as two dropdown areas:
//! - Font family list (left, wider)
//! - Font size list (right, narrower)

use tiny_skia::*;
use cosmic_text::{FontSystem, SwashCache, Color as CosmicColor};
use super::rounded_rect_path;

/// The list of available font families (matches legacy editor).
pub const FONT_FAMILIES: &[&str] = &[
    "Segoe UI",
    "Arial",
    "Helvetica",
    "Times New Roman",
    "Courier New",
    "Consolas",
    "Menlo",
    "Monaco",
    "Inter",
    "Roboto",
];

/// Available font sizes.
pub const FONT_SIZES: &[u32] = &[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

/// Result from a font picker interaction.
#[derive(Debug, Clone)]
pub enum FontPickerEvent {
    /// User picked a new font family + size.
    Changed { family: String, size: u32 },
    /// User closed without picking.
    Closed,
    /// No interaction.
    None,
}

/// The font picker widget.
pub struct FontPicker {
    /// Currently selected family.
    pub family: String,
    /// Currently selected size.
    pub size: u32,
    /// Whether the picker popup is open.
    pub open: bool,
    /// Hover index for family list.
    hover_family: Option<usize>,
    /// Hover index for size list.
    hover_size: Option<usize>,
}

const ROW_H: f32 = 22.0;
const FAMILY_W: f32 = 140.0;
const SIZE_W: f32 = 50.0;
const POPUP_GAP: f32 = 4.0;
const POPUP_PAD: f32 = 6.0;

impl FontPicker {
    pub fn new() -> Self {
        Self {
            family: "Segoe UI".to_string(),
            size: 12,
            open: false,
            hover_family: None,
            hover_size: None,
        }
    }

    pub fn set(&mut self, family: &str, size: u32) {
        self.family = family.to_string();
        self.size = size;
    }

    /// Parse "FontFamily, 12px" format.
    pub fn set_from_string(&mut self, s: &str) {
        let mut parts = s.split(',').map(|p| p.trim());
        if let Some(fam) = parts.next() {
            if !fam.is_empty() {
                self.family = fam.to_string();
            }
        }
        if let Some(size_str) = parts.next() {
            let num: String = size_str.chars().filter(|c| c.is_ascii_digit()).collect();
            if let Ok(sz) = num.parse::<u32>() {
                if sz > 0 { self.size = sz; }
            }
        }
    }

    /// Get formatted string.
    pub fn to_string(&self) -> String {
        format!("{}, {}px", self.family, self.size)
    }

    /// Popup size.
    pub fn popup_size() -> (f32, f32) {
        let family_h = FONT_FAMILIES.len() as f32 * ROW_H;
        let size_h = FONT_SIZES.len() as f32 * ROW_H;
        let h = family_h.max(size_h) + POPUP_PAD * 2.0;
        let w = POPUP_PAD * 2.0 + FAMILY_W + POPUP_GAP + SIZE_W;
        (w, h)
    }

    /// Render the compact display (current font as text, for property rows).
    pub fn render_compact(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        x: f32, y: f32, w: f32, h: f32, scale: f32,
    ) {
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Background
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(path) = rounded_rect_path(x, y, w, h, 2.0) {
            pix.fill_path(&path, &paint, FillRule::Winding, Transform::from_scale(scale, scale), None);
        }

        // Border
        paint.set_color_rgba8(180, 180, 180, 255);
        if let Some(path) = rounded_rect_path(x, y, w, h, 2.0) {
            pix.stroke_path(&path, &paint, &Stroke { width: 1.0, ..Default::default() }, Transform::from_scale(scale, scale), None);
        }

        // Text
        let display = format!("{}, {}px", self.family, self.size);
        crate::tree_view::TreeView::draw_text_static_internal(
            pix, fs, sc, &display,
            (x + 4.0) * scale, (y + 3.0) * scale,
            CosmicColor::rgba(30, 30, 30, 255), scale,
        );

        // Dropdown arrow
        let ax = x + w - 14.0;
        let ay = y + h / 2.0 - 2.0;
        paint.set_color_rgba8(80, 80, 80, 255);
        let mut pb = PathBuilder::new();
        pb.move_to(ax, ay);
        pb.line_to(ax + 8.0, ay);
        pb.line_to(ax + 4.0, ay + 5.0);
        pb.close();
        if let Some(path) = pb.finish() {
            pix.fill_path(&path, &paint, FillRule::Winding, Transform::from_scale(scale, scale), None);
        }
    }

    /// Render the popup.
    pub fn render_popup(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        popup_x: f32, popup_y: f32, scale: f32,
    ) {
        if !self.open { return; }

        let (pw, ph) = Self::popup_size();
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Shadow
        paint.set_color_rgba8(0, 0, 0, 40);
        if let Some(r) = Rect::from_xywh((popup_x + 2.0) * scale, (popup_y + 2.0) * scale, pw * scale, ph * scale) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }

        // Background
        paint.set_color_rgba8(252, 252, 252, 255);
        if let Some(path) = rounded_rect_path(popup_x, popup_y, pw, ph, 4.0) {
            pix.fill_path(&path, &paint, FillRule::Winding, Transform::from_scale(scale, scale), None);
        }

        // Border
        paint.set_color_rgba8(180, 180, 180, 255);
        if let Some(path) = rounded_rect_path(popup_x, popup_y, pw, ph, 4.0) {
            pix.stroke_path(&path, &paint, &Stroke { width: 1.0, ..Default::default() }, Transform::from_scale(scale, scale), None);
        }

        let list_x = popup_x + POPUP_PAD;
        let list_y = popup_y + POPUP_PAD;

        // ── Family list ──
        let selected_fam_idx = FONT_FAMILIES.iter().position(|f| *f == self.family.as_str());
        for (i, &fam) in FONT_FAMILIES.iter().enumerate() {
            let iy = list_y + i as f32 * ROW_H;

            // Highlight
            if Some(i) == self.hover_family {
                paint.set_color_rgba8(0, 120, 212, 40);
                if let Some(r) = Rect::from_xywh(list_x * scale, iy * scale, FAMILY_W * scale, ROW_H * scale) {
                    pix.fill_rect(r, &paint, Transform::identity(), None);
                }
            } else if Some(i) == selected_fam_idx {
                paint.set_color_rgba8(0, 120, 212, 25);
                if let Some(r) = Rect::from_xywh(list_x * scale, iy * scale, FAMILY_W * scale, ROW_H * scale) {
                    pix.fill_rect(r, &paint, Transform::identity(), None);
                }
            }

            let text_color = if Some(i) == selected_fam_idx {
                CosmicColor::rgba(0, 90, 180, 255)
            } else {
                CosmicColor::rgba(30, 30, 30, 255)
            };
            crate::tree_view::TreeView::draw_text_static_internal(
                pix, fs, sc, fam,
                (list_x + 4.0) * scale, (iy + 3.0) * scale,
                text_color, scale,
            );
        }

        // Separator line
        let sep_x = list_x + FAMILY_W + POPUP_GAP / 2.0;
        paint.set_color_rgba8(200, 200, 200, 255);
        if let Some(r) = Rect::from_xywh(sep_x * scale, list_y * scale, 1.0 * scale, (FONT_FAMILIES.len().max(FONT_SIZES.len()) as f32 * ROW_H) * scale) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }

        // ── Size list ──
        let size_x = list_x + FAMILY_W + POPUP_GAP;
        let selected_size_idx = FONT_SIZES.iter().position(|s| *s == self.size);
        for (i, &sz) in FONT_SIZES.iter().enumerate() {
            let iy = list_y + i as f32 * ROW_H;

            if Some(i) == self.hover_size {
                paint.set_color_rgba8(0, 120, 212, 40);
                if let Some(r) = Rect::from_xywh(size_x * scale, iy * scale, SIZE_W * scale, ROW_H * scale) {
                    pix.fill_rect(r, &paint, Transform::identity(), None);
                }
            } else if Some(i) == selected_size_idx {
                paint.set_color_rgba8(0, 120, 212, 25);
                if let Some(r) = Rect::from_xywh(size_x * scale, iy * scale, SIZE_W * scale, ROW_H * scale) {
                    pix.fill_rect(r, &paint, Transform::identity(), None);
                }
            }

            let text_color = if Some(i) == selected_size_idx {
                CosmicColor::rgba(0, 90, 180, 255)
            } else {
                CosmicColor::rgba(30, 30, 30, 255)
            };
            let sz_str = format!("{}", sz);
            crate::tree_view::TreeView::draw_text_static_internal(
                pix, fs, sc, &sz_str,
                (size_x + 4.0) * scale, (iy + 3.0) * scale,
                text_color, scale,
            );
        }
    }

    /// Handle a click. Returns event.
    pub fn handle_click(&mut self, mx: f32, my: f32, popup_x: f32, popup_y: f32) -> FontPickerEvent {
        if !self.open { return FontPickerEvent::None; }

        let (pw, ph) = Self::popup_size();

        // Outside → close
        if mx < popup_x || mx > popup_x + pw || my < popup_y || my > popup_y + ph {
            self.open = false;
            return FontPickerEvent::Closed;
        }

        let list_x = popup_x + POPUP_PAD;
        let list_y = popup_y + POPUP_PAD;

        // Family list click
        if mx >= list_x && mx < list_x + FAMILY_W {
            let row = ((my - list_y) / ROW_H) as usize;
            if row < FONT_FAMILIES.len() {
                self.family = FONT_FAMILIES[row].to_string();
                return FontPickerEvent::Changed { family: self.family.clone(), size: self.size };
            }
        }

        // Size list click
        let size_x = list_x + FAMILY_W + POPUP_GAP;
        if mx >= size_x && mx < size_x + SIZE_W {
            let row = ((my - list_y) / ROW_H) as usize;
            if row < FONT_SIZES.len() {
                self.size = FONT_SIZES[row];
                return FontPickerEvent::Changed { family: self.family.clone(), size: self.size };
            }
        }

        FontPickerEvent::None
    }

    /// Handle mouse hover to update highlight state.
    pub fn handle_hover(&mut self, mx: f32, my: f32, popup_x: f32, popup_y: f32) {
        if !self.open { return; }

        let list_x = popup_x + POPUP_PAD;
        let list_y = popup_y + POPUP_PAD;

        // Family hover
        if mx >= list_x && mx < list_x + FAMILY_W && my >= list_y {
            let row = ((my - list_y) / ROW_H) as usize;
            self.hover_family = if row < FONT_FAMILIES.len() { Some(row) } else { None };
        } else {
            self.hover_family = None;
        }

        // Size hover
        let size_x = list_x + FAMILY_W + POPUP_GAP;
        if mx >= size_x && mx < size_x + SIZE_W && my >= list_y {
            let row = ((my - list_y) / ROW_H) as usize;
            self.hover_size = if row < FONT_SIZES.len() { Some(row) } else { None };
        } else {
            self.hover_size = None;
        }
    }
}
