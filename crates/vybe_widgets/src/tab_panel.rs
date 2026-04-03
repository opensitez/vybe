//! TabPanel — a tab bar with switchable content panels.
//!
//! Renders tab headers at the top with an active-indicator line. Clicking
//! a tab switches the visible content panel. Only the active tab's content
//! is rendered and receives events.

use tiny_skia::*;
use cosmic_text::Color as CosmicColor;
use crate::layout::*;
use crate::ide_text;

/// One tab entry: name + content widget.
pub struct TabEntry {
    pub name: String,
    pub widget: Box<dyn PanelWidget>,
    pub closable: bool,
}

/// A tabbed panel container.
pub struct TabPanel {
    rect: LayoutRect,
    tabs: Vec<TabEntry>,
    active: usize,
    tab_height: f32,
    bg_color: (u8, u8, u8, u8),
    tab_bg: (u8, u8, u8, u8),
    tab_active_bg: (u8, u8, u8, u8),
    tab_text: (u8, u8, u8, u8),
    tab_active_text: (u8, u8, u8, u8),
    accent_color: (u8, u8, u8, u8),
}

impl TabPanel {
    pub fn new() -> Self {
        Self {
            rect: LayoutRect::zero(),
            tabs: Vec::new(),
            active: 0,
            tab_height: 28.0,
            bg_color: (30, 30, 30, 255),
            tab_bg: (45, 45, 45, 255),
            tab_active_bg: (30, 30, 30, 255),
            tab_text: (150, 150, 150, 255),
            tab_active_text: (255, 255, 255, 255),
            accent_color: (0, 122, 204, 255),
        }
    }

    pub fn set_tab_height(&mut self, h: f32) { self.tab_height = h; }

    pub fn set_colors(
        &mut self,
        bg: (u8, u8, u8, u8),
        tab_bg: (u8, u8, u8, u8),
        tab_active_bg: (u8, u8, u8, u8),
        tab_text: (u8, u8, u8, u8),
        tab_active_text: (u8, u8, u8, u8),
        accent: (u8, u8, u8, u8),
    ) {
        self.bg_color = bg;
        self.tab_bg = tab_bg;
        self.tab_active_bg = tab_active_bg;
        self.tab_text = tab_text;
        self.tab_active_text = tab_active_text;
        self.accent_color = accent;
    }

    pub fn add_tab(&mut self, name: &str, widget: Box<dyn PanelWidget>, closable: bool) {
        self.tabs.push(TabEntry {
            name: name.to_string(),
            widget,
            closable,
        });
        self.relayout();
    }

    pub fn active_index(&self) -> usize { self.active }

    pub fn set_active(&mut self, index: usize) {
        if index < self.tabs.len() {
            self.active = index;
            self.relayout();
        }
    }

    pub fn tab_count(&self) -> usize { self.tabs.len() }

    pub fn tab_name(&self, index: usize) -> Option<&str> {
        self.tabs.get(index).map(|t| t.name.as_str())
    }

    pub fn set_tab_name(&mut self, index: usize, name: &str) {
        if let Some(tab) = self.tabs.get_mut(index) {
            tab.name = name.to_string();
        }
    }

    pub fn active_widget(&self) -> Option<&dyn PanelWidget> {
        self.tabs.get(self.active).map(|t| &*t.widget)
    }

    pub fn active_widget_mut(&mut self) -> Option<&mut (dyn PanelWidget + 'static)> {
        self.tabs.get_mut(self.active).map(|t| t.widget.as_mut())
    }

    pub fn tab_widget(&self, index: usize) -> Option<&dyn PanelWidget> {
        self.tabs.get(index).map(|t| &*t.widget)
    }

    pub fn tab_widget_mut(&mut self, index: usize) -> Option<&mut (dyn PanelWidget + 'static)> {
        self.tabs.get_mut(index).map(|t| t.widget.as_mut())
    }

    pub fn remove_tab(&mut self, index: usize) -> Option<TabEntry> {
        if index < self.tabs.len() {
            let entry = self.tabs.remove(index);
            if self.active >= self.tabs.len() && self.active > 0 {
                self.active -= 1;
            }
            self.relayout();
            Some(entry)
        } else {
            None
        }
    }

    pub fn tab_height(&self) -> f32 { self.tab_height }

    fn content_rect(&self) -> LayoutRect {
        LayoutRect::new(
            self.rect.x,
            self.rect.y + self.tab_height,
            self.rect.w,
            (self.rect.h - self.tab_height).max(0.0),
        )
    }

    fn relayout(&mut self) {
        let cr = self.content_rect();
        if let Some(tab) = self.tabs.get_mut(self.active) {
            tab.widget.set_rect(cr);
        }
    }

    fn tab_width(&self) -> f32 { 120.0 }
}

impl PanelWidget for TabPanel {
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.relayout();
    }

    fn rect(&self) -> LayoutRect { self.rect }

    fn render(&mut self, ctx: &mut RenderContext) {
        let s = ctx.scale;
        let mut paint = Paint::default();

        // Tab bar background
        let (r, g, b, a) = self.tab_bg;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(
            self.rect.x * s, self.rect.y * s,
            self.rect.w * s, self.tab_height * s,
        ) {
            ctx.pixmap.fill_rect(rect, &paint, Transform::identity(), None);
        }

        // Tab headers
        let tw = self.tab_width();
        for (i, tab) in self.tabs.iter().enumerate() {
            let tx = self.rect.x + i as f32 * tw;
            let is_active = i == self.active;

            // Tab bg
            let (r, g, b, a) = if is_active { self.tab_active_bg } else { self.tab_bg };
            paint.set_color_rgba8(r, g, b, a);
            if let Some(rect) = Rect::from_xywh(tx * s, self.rect.y * s, tw * s, self.tab_height * s) {
                ctx.pixmap.fill_rect(rect, &paint, Transform::identity(), None);
            }

            // Tab text
            let (tr, tg, tb, ta) = if is_active { self.tab_active_text } else { self.tab_text };
            ide_text::draw_text(
                ctx.pixmap, ctx.font_system, ctx.swash_cache,
                &tab.name,
                tx + 8.0,
                self.rect.y + (self.tab_height - 13.0) / 2.0,
                13.0,
                CosmicColor::rgba(tr, tg, tb, ta),
                ctx.scale,
            );

            // Active indicator line
            if is_active {
                let (ar, ag, ab, aa) = self.accent_color;
                paint.set_color_rgba8(ar, ag, ab, aa);
                if let Some(rect) = Rect::from_xywh(
                    tx * s, (self.rect.y + self.tab_height - 2.0) * s,
                    tw * s, 2.0 * s,
                ) {
                    ctx.pixmap.fill_rect(rect, &paint, Transform::identity(), None);
                }
            }
        }

        // Content area background
        let cr = self.content_rect();
        let (r, g, b, a) = self.bg_color;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(cr.x * s, cr.y * s, cr.w * s, cr.h * s) {
            ctx.pixmap.fill_rect(rect, &paint, Transform::identity(), None);
        }

        // Active tab content
        if let Some(tab) = self.tabs.get_mut(self.active) {
            tab.widget.render(ctx);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        // Tab bar click
        let tab_bar = LayoutRect::new(self.rect.x, self.rect.y, self.rect.w, self.tab_height);
        if tab_bar.contains(event.x, event.y) {
            if let MouseEventKind::Press(MouseButton::Left) = event.kind {
                let tw = self.tab_width();
                let rel_x = event.x - self.rect.x;
                let idx = (rel_x / tw) as usize;
                if idx < self.tabs.len() && idx != self.active {
                    self.active = idx;
                    self.relayout();
                }
                return true;
            }
        }

        // Route to active content
        if let Some(tab) = self.tabs.get_mut(self.active) {
            if tab.widget.rect().contains(event.x, event.y) {
                return tab.widget.handle_mouse(event);
            }
        }
        false
    }

    fn handle_key(&mut self, event: &KeyEvent) -> bool {
        if let Some(tab) = self.tabs.get_mut(self.active) {
            return tab.widget.handle_key(event);
        }
        false
    }
}
