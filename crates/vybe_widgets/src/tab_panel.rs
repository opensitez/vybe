//! TabPanel — a tab bar with switchable content panels.
//!
//! Renders tab headers at the top with an active-indicator line. Clicking
//! a tab switches the visible content panel. Only the active tab's content
//! is rendered and receives events.

use crate::ide_text;
use crate::layout::*;
use cosmic_text::Color as CosmicColor;
use tiny_skia::*;

/// One tab entry: name + content widget.
pub struct TabEntry {
    pub name: String,
    pub widget: Box<dyn PanelWidget>,
    pub closable: bool,
}

/// A tabbed panel container.
pub struct TabPanel {
    pub id: WidgetId,
    rect: LayoutRect,
    tabs: Vec<TabEntry>,
    active: usize,
    tab_height: f32,
    custom_tab_width: Option<f32>,
    bg_color: (u8, u8, u8, u8),
    tab_bg: (u8, u8, u8, u8),
    tab_active_bg: (u8, u8, u8, u8),
    tab_text: (u8, u8, u8, u8),
    tab_active_text: (u8, u8, u8, u8),
    accent_color: (u8, u8, u8, u8),
    pending_events: Vec<WidgetEvent>,
    hovering_close: Option<usize>,
    scroll_x: f32,
}

impl TabPanel {
    pub fn new() -> Self {
        Self {
            id: WidgetId::next(),
            rect: LayoutRect::zero(),
            tabs: Vec::new(),
            active: 0,
            tab_height: 28.0,
            custom_tab_width: None,
            bg_color: (30, 30, 30, 255),
            tab_bg: (45, 45, 45, 255),
            tab_active_bg: (30, 30, 30, 255),
            tab_text: (150, 150, 150, 255),
            tab_active_text: (255, 255, 255, 255),
            accent_color: (0, 122, 204, 255),
            pending_events: Vec::new(),
            hovering_close: None,
            scroll_x: 0.0,
        }
    }

    pub fn set_tab_height(&mut self, h: f32) {
        self.tab_height = h;
    }

    pub fn set_tab_width(&mut self, w: f32) {
        self.custom_tab_width = Some(w);
    }

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

    /// Add a tab header with no content widget (uses NullWidget).
    /// The host app manages content rendering separately.
    pub fn add_tab_header(&mut self, name: &str, closable: bool) {
        self.tabs.push(TabEntry {
            name: name.to_string(),
            widget: Box::new(NullWidget::new()),
            closable,
        });
    }

    /// Insert a tab header at a specific position.
    pub fn insert_tab_header(&mut self, index: usize, name: &str, closable: bool) {
        let idx = index.min(self.tabs.len());
        self.tabs.insert(
            idx,
            TabEntry {
                name: name.to_string(),
                widget: Box::new(NullWidget::new()),
                closable,
            },
        );
    }

    pub fn active_index(&self) -> usize {
        self.active
    }

    pub fn set_active(&mut self, index: usize) {
        if index < self.tabs.len() {
            self.active = index;
            self.relayout();
        }
    }

    pub fn tab_count(&self) -> usize {
        self.tabs.len()
    }

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

    pub fn tab_height(&self) -> f32 {
        self.tab_height
    }

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

    fn tab_width(&self) -> f32 {
        self.custom_tab_width.unwrap_or(120.0)
    }

    /// Scroll the tab bar (e.g. from mouse wheel over tab bar).
    pub fn scroll_tab_bar(&mut self, delta: f32) {
        self.scroll_x = (self.scroll_x - delta).max(0.0);
        let max_scroll = (self.tabs.len() as f32 * self.tab_width() - self.rect.w).max(0.0);
        self.scroll_x = self.scroll_x.min(max_scroll);
    }
}

impl PanelWidget for TabPanel {
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.relayout();
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

        // Tab bar background
        let (r, g, b, a) = self.tab_bg;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(
            self.rect.x * s,
            self.rect.y * s,
            self.rect.w * s,
            self.tab_height * s,
        ) {
            ctx.pixmap
                .fill_rect(rect, &paint, Transform::identity(), None);
        }

        // Tab headers
        let tw = self.tab_width();
        for (i, tab) in self.tabs.iter().enumerate() {
            let tx = self.rect.x + i as f32 * tw - self.scroll_x;
            // Skip tabs that are scrolled out of view
            if tx + tw < self.rect.x || tx > self.rect.right() {
                continue;
            }
            let is_active = i == self.active;

            // Tab bg
            let (r, g, b, a) = if is_active {
                self.tab_active_bg
            } else {
                self.tab_bg
            };
            paint.set_color_rgba8(r, g, b, a);
            if let Some(rect) =
                Rect::from_xywh(tx * s, self.rect.y * s, tw * s, self.tab_height * s)
            {
                ctx.pixmap
                    .fill_rect(rect, &paint, Transform::identity(), None);
            }

            // Tab text
            let (tr, tg, tb, ta) = if is_active {
                self.tab_active_text
            } else {
                self.tab_text
            };
            ide_text::draw_text(
                ctx.pixmap,
                ctx.font_system,
                ctx.swash_cache,
                &tab.name,
                tx + 8.0,
                self.rect.y + (self.tab_height - 13.0) / 2.0,
                13.0,
                CosmicColor::rgba(tr, tg, tb, ta),
                ctx.scale,
            );

            // Close button (×) for closable tabs
            if tab.closable {
                let close_x = tx + tw - 20.0;
                let close_y = self.rect.y + (self.tab_height - 12.0) / 2.0;
                let hover = self.hovering_close == Some(i);
                let close_col = if hover {
                    CosmicColor::rgba(255, 100, 100, 255)
                } else {
                    CosmicColor::rgba(160, 160, 160, 255)
                };
                ide_text::draw_text(
                    ctx.pixmap,
                    ctx.font_system,
                    ctx.swash_cache,
                    "×",
                    close_x,
                    close_y,
                    12.0,
                    close_col,
                    ctx.scale,
                );
            }

            // Active indicator line
            if is_active {
                let (ar, ag, ab, aa) = self.accent_color;
                paint.set_color_rgba8(ar, ag, ab, aa);
                if let Some(rect) = Rect::from_xywh(
                    tx * s,
                    (self.rect.y + self.tab_height - 2.0) * s,
                    tw * s,
                    2.0 * s,
                ) {
                    ctx.pixmap
                        .fill_rect(rect, &paint, Transform::identity(), None);
                }
            }
        }

        // Content area background
        let cr = self.content_rect();
        let (r, g, b, a) = self.bg_color;
        paint.set_color_rgba8(r, g, b, a);
        if let Some(rect) = Rect::from_xywh(cr.x * s, cr.y * s, cr.w * s, cr.h * s) {
            ctx.pixmap
                .fill_rect(rect, &paint, Transform::identity(), None);
        }

        // Active tab content
        if let Some(tab) = self.tabs.get_mut(self.active) {
            tab.widget.render(ctx);
        }
    }

    fn handle_mouse(&mut self, event: &MouseEvent) -> bool {
        // Tab bar region
        let tab_bar = LayoutRect::new(self.rect.x, self.rect.y, self.rect.w, self.tab_height);
        if tab_bar.contains(event.x, event.y) {
            let tw = self.tab_width();
            let rel_x = event.x - self.rect.x + self.scroll_x;
            let idx = (rel_x / tw) as usize;

            match event.kind {
                MouseEventKind::Press(MouseButton::Left) => {
                    if idx < self.tabs.len() {
                        // Check close button hit (last 20px of tab)
                        let tab_local_x = rel_x - idx as f32 * tw;
                        if self.tabs[idx].closable && tab_local_x >= tw - 22.0 {
                            self.pending_events
                                .push(WidgetEvent::TabCloseRequested(idx));
                        } else if idx != self.active {
                            self.active = idx;
                            self.relayout();
                            self.pending_events.push(WidgetEvent::TabChanged(idx));
                        }
                    }
                    return true;
                }
                MouseEventKind::Move => {
                    // Update close button hover
                    let old = self.hovering_close;
                    self.hovering_close = None;
                    if idx < self.tabs.len() && self.tabs[idx].closable {
                        let tab_local_x = rel_x - idx as f32 * tw;
                        if tab_local_x >= tw - 22.0 {
                            self.hovering_close = Some(idx);
                        }
                    }
                    return self.hovering_close != old;
                }
                _ => {}
            }
            return false;
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

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        if let Some(tab) = self.tabs.get_mut(self.active) {
            if tab.widget.rect().contains(x, y) {
                return tab.widget.handle_scroll(delta, x, y);
            }
        }
        false
    }

    fn cursor_at(&self, x: f32, y: f32) -> winit::window::CursorIcon {
        if let Some(tab) = self.tabs.get(self.active) {
            if tab.widget.rect().contains(x, y) {
                return tab.widget.cursor_at(x, y);
            }
        }
        winit::window::CursorIcon::Default
    }

    fn drain_events(&mut self) -> Vec<WidgetEvent> {
        let mut events = Vec::new();
        events.extend(self.pending_events.drain(..));
        for tab in &mut self.tabs {
            events.extend(tab.widget.drain_events());
        }
        events
    }
}
