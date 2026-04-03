//! Color Picker widget — standalone tiny-skia rendered color picker.
//!
//! Features:
//! - Hue bar (vertical strip on the right)
//! - Saturation/Value square (main area)
//! - Current color swatch + hex display
//! - Click to select color

use tiny_skia::*;
use cosmic_text::Color as CosmicColor;
use super::{WidgetColors, rounded_rect_path};
use super::layout::{
    LayoutRect, MouseEvent, MouseEventKind, MouseButton as LayoutMouseButton,
    KeyEvent, RenderContext, PanelWidget, WidgetEvent,
};

/// A picked color in RGBA.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PickedColor {
    pub r: u8,
    pub g: u8,
    pub b: u8,
    pub a: u8,
}

impl PickedColor {
    pub fn from_rgba(r: u8, g: u8, b: u8, a: u8) -> Self {
        Self { r, g, b, a }
    }

    pub fn to_hex(&self) -> String {
        format!("#{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }

    pub fn from_hex(hex: &str) -> Option<Self> {
        let hex = hex.trim_start_matches('#');
        if hex.len() < 6 { return None; }
        let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
        Some(Self { r, g, b, a: 255 })
    }
}

impl Default for PickedColor {
    fn default() -> Self {
        Self { r: 255, g: 255, b: 255, a: 255 }
    }
}

/// Result from a color picker interaction.
#[derive(Debug)]
pub enum ColorPickerEvent {
    /// User picked a color.
    Changed(PickedColor),
    /// User closed without picking.
    Closed,
    /// No interaction.
    None,
}

/// HSV color for internal calculations.
#[derive(Clone, Copy, Debug)]
struct Hsv {
    h: f32, // 0..360
    s: f32, // 0..1
    v: f32, // 0..1
}

impl Hsv {
    fn to_rgb(&self) -> (u8, u8, u8) {
        let h = self.h;
        let s = self.s;
        let v = self.v;
        let c = v * s;
        let x = c * (1.0 - ((h / 60.0) % 2.0 - 1.0).abs());
        let m = v - c;
        let (r1, g1, b1) = if h < 60.0 {
            (c, x, 0.0)
        } else if h < 120.0 {
            (x, c, 0.0)
        } else if h < 180.0 {
            (0.0, c, x)
        } else if h < 240.0 {
            (0.0, x, c)
        } else if h < 300.0 {
            (x, 0.0, c)
        } else {
            (c, 0.0, x)
        };
        (
            ((r1 + m) * 255.0) as u8,
            ((g1 + m) * 255.0) as u8,
            ((b1 + m) * 255.0) as u8,
        )
    }

    fn from_rgb(r: u8, g: u8, b: u8) -> Self {
        let rf = r as f32 / 255.0;
        let gf = g as f32 / 255.0;
        let bf = b as f32 / 255.0;
        let max = rf.max(gf).max(bf);
        let min = rf.min(gf).min(bf);
        let delta = max - min;
        let h = if delta == 0.0 {
            0.0
        } else if max == rf {
            60.0 * (((gf - bf) / delta) % 6.0)
        } else if max == gf {
            60.0 * (((bf - rf) / delta) + 2.0)
        } else {
            60.0 * (((rf - gf) / delta) + 4.0)
        };
        let h = if h < 0.0 { h + 360.0 } else { h };
        let s = if max == 0.0 { 0.0 } else { delta / max };
        let v = max;
        Self { h, s, v }
    }
}

/// The color picker widget.
pub struct ColorPicker {
    /// Current HSV state.
    hsv: Hsv,
    /// Current picked color.
    pub color: PickedColor,
    /// Whether the picker popup is visible.
    pub open: bool,
    /// Dragging in the SV square.
    dragging_sv: bool,
    /// Dragging in the hue bar.
    dragging_hue: bool,
    /// Widget name for events.
    pub name: String,
    /// Layout rect for PanelWidget.
    rect: LayoutRect,
    /// Pending events.
    pending_events: Vec<WidgetEvent>,
}

/// Layout constants (in logical pixels).
const SV_SIZE: f32 = 150.0;
const HUE_W: f32 = 18.0;
const GAP: f32 = 8.0;
const SWATCH_H: f32 = 24.0;
const PADDING: f32 = 8.0;

impl ColorPicker {
    pub fn new() -> Self {
        Self {
            hsv: Hsv { h: 0.0, s: 1.0, v: 1.0 },
            color: PickedColor::default(),
            open: false,
            dragging_sv: false,
            dragging_hue: false,
            name: String::new(),
            rect: LayoutRect::zero(),
            pending_events: Vec::new(),
        }
    }

    pub fn with_name(mut self, name: &str) -> Self { self.name = name.to_string(); self }

    pub fn set_color(&mut self, c: PickedColor) {
        self.color = c;
        self.hsv = Hsv::from_rgb(c.r, c.g, c.b);
    }

    pub fn set_from_hex(&mut self, hex: &str) {
        if let Some(c) = PickedColor::from_hex(hex) {
            self.set_color(c);
        }
    }

    /// Total popup size.
    pub fn popup_size() -> (f32, f32) {
        let w = PADDING * 2.0 + SV_SIZE + GAP + HUE_W;
        let h = PADDING * 2.0 + SV_SIZE + GAP + SWATCH_H;
        (w, h)
    }

    /// Render the small swatch button (used inline in property rows).
    pub fn render_swatch(&self, pix: &mut Pixmap, x: f32, y: f32, w: f32, h: f32, scale: f32) {
        let mut paint = Paint::default();
        paint.anti_alias = true;

        // Checkerboard background (for alpha)
        paint.set_color_rgba8(200, 200, 200, 255);
        if let Some(r) = Rect::from_xywh(x * scale, y * scale, w * scale, h * scale) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }
        // White squares
        paint.set_color_rgba8(255, 255, 255, 255);
        let half_w = w / 2.0;
        let half_h = h / 2.0;
        if let Some(r) = Rect::from_xywh(x * scale, y * scale, half_w * scale, half_h * scale) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }
        if let Some(r) = Rect::from_xywh((x + half_w) * scale, (y + half_h) * scale, half_w * scale, half_h * scale) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }

        // Actual color
        paint.set_color_rgba8(self.color.r, self.color.g, self.color.b, self.color.a);
        if let Some(path) = rounded_rect_path(x, y, w, h, 2.0) {
            pix.fill_path(&path, &paint, FillRule::Winding, Transform::from_scale(scale, scale), None);
        }

        // Border
        paint.set_color_rgba8(118, 118, 118, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        if let Some(path) = rounded_rect_path(x, y, w, h, 2.0) {
            pix.stroke_path(&path, &paint, &stroke, Transform::from_scale(scale, scale), None);
        }
    }

    /// Render the full color picker popup.
    pub fn render_popup(&self, pix: &mut Pixmap, popup_x: f32, popup_y: f32, scale: f32) {
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
        paint.set_color_rgba8(250, 250, 250, 255);
        if let Some(path) = rounded_rect_path(popup_x, popup_y, pw, ph, 4.0) {
            pix.fill_path(&path, &paint, FillRule::Winding, Transform::from_scale(scale, scale), None);
        }

        // Border
        paint.set_color_rgba8(180, 180, 180, 255);
        let mut stroke = Stroke::default();
        stroke.width = 1.0;
        if let Some(path) = rounded_rect_path(popup_x, popup_y, pw, ph, 4.0) {
            pix.stroke_path(&path, &paint, &stroke, Transform::from_scale(scale, scale), None);
        }

        let sv_x = popup_x + PADDING;
        let sv_y = popup_y + PADDING;

        // ── Saturation/Value square ──
        // Render pixel-by-pixel (at reduced resolution for performance)
        let step = 3.0; // pixels per sample
        let mut sx = 0.0_f32;
        while sx < SV_SIZE {
            let mut sy = 0.0_f32;
            while sy < SV_SIZE {
                let s = sx / SV_SIZE;
                let v = 1.0 - (sy / SV_SIZE);
                let hsv = Hsv { h: self.hsv.h, s, v };
                let (r, g, b) = hsv.to_rgb();
                paint.set_color_rgba8(r, g, b, 255);
                if let Some(r) = Rect::from_xywh(
                    (sv_x + sx) * scale,
                    (sv_y + sy) * scale,
                    step * scale,
                    step * scale,
                ) {
                    pix.fill_rect(r, &paint, Transform::identity(), None);
                }
                sy += step;
            }
            sx += step;
        }

        // SV square border
        paint.set_color_rgba8(160, 160, 160, 255);
        let mut pb = PathBuilder::new();
        if let Some(r) = Rect::from_xywh(sv_x * scale, sv_y * scale, SV_SIZE * scale, SV_SIZE * scale) {
            pb.push_rect(r);
        }
        if let Some(path) = pb.finish() {
            pix.stroke_path(&path, &paint, &Stroke { width: 1.0 * scale, ..Default::default() }, Transform::identity(), None);
        }

        // SV cursor (circle)
        let cx = sv_x + self.hsv.s * SV_SIZE;
        let cy = sv_y + (1.0 - self.hsv.v) * SV_SIZE;
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(path) = super::circle_path(cx, cy, 5.0) {
            pix.stroke_path(&path, &paint, &Stroke { width: 2.0, ..Default::default() }, Transform::from_scale(scale, scale), None);
        }
        paint.set_color_rgba8(0, 0, 0, 255);
        if let Some(path) = super::circle_path(cx, cy, 4.0) {
            pix.stroke_path(&path, &paint, &Stroke { width: 1.0, ..Default::default() }, Transform::from_scale(scale, scale), None);
        }

        // ── Hue bar ──
        let hue_x = sv_x + SV_SIZE + GAP;
        let hue_y = sv_y;
        let hue_h = SV_SIZE;
        let hue_step = 2.0;
        let mut hy = 0.0_f32;
        while hy < hue_h {
            let h = (hy / hue_h) * 360.0;
            let hsv = Hsv { h, s: 1.0, v: 1.0 };
            let (r, g, b) = hsv.to_rgb();
            paint.set_color_rgba8(r, g, b, 255);
            if let Some(r) = Rect::from_xywh(
                hue_x * scale,
                (hue_y + hy) * scale,
                HUE_W * scale,
                hue_step * scale,
            ) {
                pix.fill_rect(r, &paint, Transform::identity(), None);
            }
            hy += hue_step;
        }

        // Hue bar border
        paint.set_color_rgba8(160, 160, 160, 255);
        let mut pb = PathBuilder::new();
        if let Some(r) = Rect::from_xywh(hue_x * scale, hue_y * scale, HUE_W * scale, hue_h * scale) {
            pb.push_rect(r);
        }
        if let Some(path) = pb.finish() {
            pix.stroke_path(&path, &paint, &Stroke { width: 1.0 * scale, ..Default::default() }, Transform::identity(), None);
        }

        // Hue cursor (two horizontal lines)
        let hue_cursor_y = hue_y + (self.hsv.h / 360.0) * hue_h;
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(r) = Rect::from_xywh((hue_x - 1.0) * scale, (hue_cursor_y - 1.5) * scale, (HUE_W + 2.0) * scale, 3.0 * scale) {
            pix.fill_rect(r, &paint, Transform::identity(), None);
        }
        paint.set_color_rgba8(0, 0, 0, 255);
        let mut pb = PathBuilder::new();
        if let Some(r) = Rect::from_xywh((hue_x - 1.0) * scale, (hue_cursor_y - 1.5) * scale, (HUE_W + 2.0) * scale, 3.0 * scale) {
            pb.push_rect(r);
        }
        if let Some(path) = pb.finish() {
            pix.stroke_path(&path, &paint, &Stroke { width: 1.0 * scale, ..Default::default() }, Transform::identity(), None);
        }

        // ── Swatch + hex text at bottom ──
        let swatch_y = sv_y + SV_SIZE + GAP;
        // Color swatch
        paint.set_color_rgba8(self.color.r, self.color.g, self.color.b, self.color.a);
        if let Some(path) = rounded_rect_path(sv_x, swatch_y, SWATCH_H, SWATCH_H, 3.0) {
            pix.fill_path(&path, &paint, FillRule::Winding, Transform::from_scale(scale, scale), None);
        }
        paint.set_color_rgba8(160, 160, 160, 255);
        if let Some(path) = rounded_rect_path(sv_x, swatch_y, SWATCH_H, SWATCH_H, 3.0) {
            pix.stroke_path(&path, &paint, &Stroke { width: 1.0, ..Default::default() }, Transform::from_scale(scale, scale), None);
        }
    }

    /// Handle a click at (mx, my) relative to the popup position.
    /// Returns a `ColorPickerEvent`.
    pub fn handle_click(&mut self, mx: f32, my: f32, popup_x: f32, popup_y: f32) -> ColorPickerEvent {
        if !self.open { return ColorPickerEvent::None; }

        let (pw, ph) = Self::popup_size();

        // Outside popup → close
        if mx < popup_x || mx > popup_x + pw || my < popup_y || my > popup_y + ph {
            self.open = false;
            self.dragging_sv = false;
            self.dragging_hue = false;
            return ColorPickerEvent::Closed;
        }

        let sv_x = popup_x + PADDING;
        let sv_y = popup_y + PADDING;
        let hue_x = sv_x + SV_SIZE + GAP;
        let hue_y = sv_y;

        // Click in SV square
        if mx >= sv_x && mx <= sv_x + SV_SIZE && my >= sv_y && my <= sv_y + SV_SIZE {
            self.dragging_sv = true;
            self.hsv.s = ((mx - sv_x) / SV_SIZE).clamp(0.0, 1.0);
            self.hsv.v = 1.0 - ((my - sv_y) / SV_SIZE).clamp(0.0, 1.0);
            let (r, g, b) = self.hsv.to_rgb();
            self.color = PickedColor::from_rgba(r, g, b, 255);
            return ColorPickerEvent::Changed(self.color);
        }

        // Click in hue bar
        if mx >= hue_x && mx <= hue_x + HUE_W && my >= hue_y && my <= hue_y + SV_SIZE {
            self.dragging_hue = true;
            self.hsv.h = (((my - hue_y) / SV_SIZE) * 360.0).clamp(0.0, 359.99);
            let (r, g, b) = self.hsv.to_rgb();
            self.color = PickedColor::from_rgba(r, g, b, 255);
            return ColorPickerEvent::Changed(self.color);
        }

        ColorPickerEvent::None
    }

    /// Handle mouse drag (call on mouse move while button is down).
    pub fn handle_drag(&mut self, mx: f32, my: f32, popup_x: f32, popup_y: f32) -> ColorPickerEvent {
        if !self.open { return ColorPickerEvent::None; }

        let sv_x = popup_x + PADDING;
        let sv_y = popup_y + PADDING;
        let hue_x = sv_x + SV_SIZE + GAP;
        let hue_y = sv_y;

        if self.dragging_sv {
            self.hsv.s = ((mx - sv_x) / SV_SIZE).clamp(0.0, 1.0);
            self.hsv.v = 1.0 - ((my - sv_y) / SV_SIZE).clamp(0.0, 1.0);
            let (r, g, b) = self.hsv.to_rgb();
            self.color = PickedColor::from_rgba(r, g, b, 255);
            return ColorPickerEvent::Changed(self.color);
        }

        if self.dragging_hue {
            self.hsv.h = (((my - hue_y) / SV_SIZE) * 360.0).clamp(0.0, 359.99);
            let (r, g, b) = self.hsv.to_rgb();
            self.color = PickedColor::from_rgba(r, g, b, 255);
            return ColorPickerEvent::Changed(self.color);
        }

        ColorPickerEvent::None
    }

    /// Call on mouse up to stop dragging.
    pub fn handle_mouse_up(&mut self) {
        self.dragging_sv = false;
        self.dragging_hue = false;
    }

    /// Internal: compute SV square and hue bar positions relative to a given origin.
    fn layout_regions(&self, ox: f32, oy: f32) -> (f32, f32, f32, f32) {
        let sv_x = ox + PADDING;
        let sv_y = oy + PADDING;
        let hue_x = sv_x + SV_SIZE + GAP;
        let hue_y = sv_y;
        (sv_x, sv_y, hue_x, hue_y)
    }

    /// Internal: update HSV from an SV-square click and emit event.
    fn pick_sv(&mut self, mx: f32, my: f32, sv_x: f32, sv_y: f32) {
        self.hsv.s = ((mx - sv_x) / SV_SIZE).clamp(0.0, 1.0);
        self.hsv.v = 1.0 - ((my - sv_y) / SV_SIZE).clamp(0.0, 1.0);
        let (r, g, b) = self.hsv.to_rgb();
        self.color = PickedColor::from_rgba(r, g, b, 255);
        self.pending_events.push(WidgetEvent::ColorChanged(
            self.name.clone(),
            self.color.to_hex(),
        ));
    }

    /// Internal: update hue from a hue-bar click and emit event.
    fn pick_hue(&mut self, my: f32, hue_y: f32) {
        self.hsv.h = (((my - hue_y) / SV_SIZE) * 360.0).clamp(0.0, 359.99);
        let (r, g, b) = self.hsv.to_rgb();
        self.color = PickedColor::from_rgba(r, g, b, 255);
        self.pending_events.push(WidgetEvent::ColorChanged(
            self.name.clone(),
            self.color.to_hex(),
        ));
    }
}

// ── PanelWidget impl ───────────────────────────────────────────────────

impl PanelWidget for ColorPicker {
    fn name(&self) -> &str { &self.name }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
    }

    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 { return; }

        // Render the full picker inline — temporarily set open so render_popup works
        let was_open = self.open;
        self.open = true;
        self.render_popup(ctx.pixmap, r.x, r.y, ctx.scale);
        self.open = was_open;

        // Also draw hex text next to the swatch at the bottom
        let sv_y = r.y + PADDING;
        let swatch_y = sv_y + SV_SIZE + GAP;
        let text_x = r.x + PADDING + SWATCH_H + 6.0;
        let hex = self.color.to_hex();
        super::ide_text::draw_text(
            ctx.pixmap, ctx.font_system, ctx.swash_cache,
            &hex, text_x, swatch_y + 4.0, 12.0,
            CosmicColor::rgba(60, 60, 60, 255), ctx.scale,
        );
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) && !self.dragging_sv && !self.dragging_hue {
            return false;
        }

        let (sv_x, sv_y, hue_x, hue_y) = self.layout_regions(self.rect.x, self.rect.y);

        match event.kind {
            MouseEventKind::Press(LayoutMouseButton::Left) => {
                // SV square hit
                if event.x >= sv_x && event.x <= sv_x + SV_SIZE
                    && event.y >= sv_y && event.y <= sv_y + SV_SIZE
                {
                    self.dragging_sv = true;
                    self.pick_sv(event.x, event.y, sv_x, sv_y);
                    return true;
                }
                // Hue bar hit
                if event.x >= hue_x && event.x <= hue_x + HUE_W
                    && event.y >= hue_y && event.y <= hue_y + SV_SIZE
                {
                    self.dragging_hue = true;
                    self.pick_hue(event.y, hue_y);
                    return true;
                }
                // Clicked inside widget rect but not on a control area
                true
            }
            MouseEventKind::Move => {
                if self.dragging_sv {
                    self.pick_sv(event.x, event.y, sv_x, sv_y);
                    return true;
                }
                if self.dragging_hue {
                    self.pick_hue(event.y, hue_y);
                    return true;
                }
                false
            }
            MouseEventKind::Release(LayoutMouseButton::Left) => {
                let was_dragging = self.dragging_sv || self.dragging_hue;
                self.dragging_sv = false;
                self.dragging_hue = false;
                was_dragging
            }
            _ => false,
        }
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool { false }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }
}
