//! StatusBarPanel — a bottom status bar with left- and right-aligned sections.
//!
//! Sections can hold text labels (e.g. line/col, language, build config).
//! Each section can optionally have a click identifier for the host to handle.

use crate::ide_text;
use crate::layout::*;
use cosmic_text::Color as CosmicColor;
use tiny_skia::*;

/// A single section in the status bar.
pub struct StatusSection {
    pub text: String,
    pub width: f32,
    pub align_right: bool,
    pub fg: (u8, u8, u8, u8),
    pub click_id: Option<String>,
}

/// A bottom status bar with left- and right-aligned sections.
pub struct StatusBarPanel {
    pub id: WidgetId,
    rect: LayoutRect,
    sections: Vec<StatusSection>,
    height: f32,
    bg_color: (u8, u8, u8, u8),
    pending_events: Vec<WidgetEvent>,
}

impl StatusBarPanel {
    pub fn new() -> Self {
        Self {
            id: WidgetId::next(),
            rect: LayoutRect::zero(),
            sections: Vec::new(),
            height: 24.0,
            bg_color: (0, 122, 204, 255),
            pending_events: Vec::new(),
        }
    }

    pub fn set_background(&mut self, r: u8, g: u8, b: u8, a: u8) {
        self.bg_color = (r, g, b, a);
    }

    pub fn set_height(&mut self, h: f32) {
        self.height = h;
    }
    pub fn height(&self) -> f32 {
        self.height
    }

    pub fn add_section(&mut self, text: &str, width: f32, align_right: bool) {
        self.sections.push(StatusSection {
            text: text.to_string(),
            width,
            align_right,
            fg: (255, 255, 255, 255),
            click_id: None,
        });
    }

    pub fn add_section_with_id(
        &mut self,
        text: &str,
        width: f32,
        align_right: bool,
        click_id: &str,
    ) {
        self.sections.push(StatusSection {
            text: text.to_string(),
            width,
            align_right,
            fg: (255, 255, 255, 255),
            click_id: Some(click_id.to_string()),
        });
    }

    pub fn set_section_text(&mut self, index: usize, text: &str) {
        if let Some(sec) = self.sections.get_mut(index) {
            sec.text = text.to_string();
        }
    }

    pub fn set_section_fg(&mut self, index: usize, r: u8, g: u8, b: u8, a: u8) {
        if let Some(sec) = self.sections.get_mut(index) {
            sec.fg = (r, g, b, a);
        }
    }

    pub fn sections(&self) -> &[StatusSection] {
        &self.sections
    }
    pub fn sections_mut(&mut self) -> &mut Vec<StatusSection> {
        &mut self.sections
    }

    /// Returns the `click_id` of the section at the given point, if any.
    pub fn hit_test_section(&self, x: f32, y: f32) -> Option<&str> {
        if !self.rect.contains(x, y) {
            return None;
        }

        let mut lx = self.rect.x + 8.0;
        for sec in &self.sections {
            if sec.align_right {
                continue;
            }
            let sec_rect = LayoutRect::new(lx, self.rect.y, sec.width, self.height);
            if sec_rect.contains(x, y) {
                return sec.click_id.as_deref();
            }
            lx += sec.width + 8.0;
        }

        let mut rx = self.rect.right() - 8.0;
        for sec in self.sections.iter().rev() {
            if !sec.align_right {
                continue;
            }
            rx -= sec.width;
            let sec_rect = LayoutRect::new(rx, self.rect.y, sec.width, self.height);
            if sec_rect.contains(x, y) {
                return sec.click_id.as_deref();
            }
            rx -= 8.0;
        }

        None
    }
}

impl PanelWidget for StatusBarPanel {
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = LayoutRect::new(rect.x, rect.y, rect.w, self.height);
    }

    fn rect(&self) -> LayoutRect {
        self.rect
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        let s = ctx.scale;
        let mut paint = Paint::default();
        let (r, g, b, a) = self.bg_color;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(
            self.rect.x * s,
            self.rect.y * s,
            self.rect.w * s,
            self.rect.h * s,
        ) {
            ctx.pixmap
                .fill_rect(rect, &paint, Transform::identity(), None);
        }

        let text_y = self.rect.y + (self.height - 12.0) / 2.0;

        // Left-aligned sections
        let mut lx = self.rect.x + 8.0;
        for sec in &self.sections {
            if sec.align_right {
                continue;
            }
            let (fr, fg, fb, fa) = sec.fg;
            ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &sec.text,
                lx,
                text_y,
                12.0,
                CosmicColor::rgba(fr, fg, fb, fa),
                ctx.scale,
            );
            lx += sec.width + 8.0;
        }

        // Right-aligned sections (drawn right to left)
        let mut rx = self.rect.right() - 8.0;
        for sec in self.sections.iter().rev() {
            if !sec.align_right {
                continue;
            }
            rx -= sec.width;
            let (fr, fg, fb, fa) = sec.fg;
            ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &sec.text,
                rx,
                text_y,
                12.0,
                CosmicColor::rgba(fr, fg, fb, fa),
                ctx.scale,
            );
            rx -= 8.0;
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        if !self.rect.contains(event.x, event.y) {
            return false;
        }
        if let MouseEventKind::Press(_) = &event.kind {
            if let Some(id) = self.hit_test_section(event.x, event.y) {
                self.pending_events
                    .push(WidgetEvent::StatusBarClick(id.to_string()));
            }
            return true;
        }
        false
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        std::mem::take(&mut self.pending_events)
    }

    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }
}
