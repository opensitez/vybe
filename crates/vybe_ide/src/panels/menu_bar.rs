//! Menu bar panel — File, Edit, View, Project, Build, Help menus.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform};

use crate::layout::Rect;
use crate::text::draw_text;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MenuAction {
    NewProject,
    OpenProject,
    SaveProject,
    SaveAs,
    Exit,
    Undo,
    Redo,
    Cut,
    Copy,
    Paste,
    AddForm,
    AddModule,
    RunProject,
    StopProject,
    About,
}

const MENUS: &[&str] = &["File", "Edit", "Project", "Run"];

const FILE_ITEMS: &[(&str, MenuAction)] = &[
    ("New Project", MenuAction::NewProject),
    ("Open Project...", MenuAction::OpenProject),
    ("Save Project", MenuAction::SaveProject),
    ("Save Project As...", MenuAction::SaveAs),
    ("Exit", MenuAction::Exit),
];

const EDIT_ITEMS: &[(&str, MenuAction)] = &[
    ("Undo          Ctrl+Z", MenuAction::Undo),
    ("Redo          Ctrl+Y", MenuAction::Redo),
    ("Delete", MenuAction::Cut),
    ("Cut", MenuAction::Cut),
    ("Copy", MenuAction::Copy),
    ("Paste", MenuAction::Paste),
];

const PROJECT_ITEMS: &[(&str, MenuAction)] = &[
    ("Add Form", MenuAction::AddForm),
    ("Add Module", MenuAction::AddModule),
    ("Project Properties...", MenuAction::About),
];

const RUN_ITEMS: &[(&str, MenuAction)] = &[
    ("Start", MenuAction::RunProject),
    ("End", MenuAction::StopProject),
];

/// Logical width of a menu label.
fn menu_label_w(label: &str) -> f32 {
    label.len() as f32 * 9.0 + 20.0
}

pub struct MenuBar {
    pub open_menu: Option<usize>,
    pub hover_item: Option<usize>,
    pub hover_menu: Option<usize>,
}

impl MenuBar {
    pub fn new() -> Self {
        Self { open_menu: None, hover_item: None, hover_menu: None }
    }

    /// Call on mouse move to update hover state.
    pub fn handle_hover(&mut self, mx: f32, my: f32, rect: Rect) {
        self.hover_item = None;
        self.hover_menu = None;

        // Hovering over menu labels?
        if my >= rect.y && my < rect.y + rect.h {
            let mut x = rect.x + 6.0;
            for (i, label) in MENUS.iter().enumerate() {
                let lw = menu_label_w(label);
                if mx >= x && mx < x + lw {
                    self.hover_menu = Some(i);
                    // If a menu is already open, switch to this one on hover
                    if self.open_menu.is_some() {
                        self.open_menu = Some(i);
                    }
                    return;
                }
                x += lw;
            }
            return;
        }

        // Hovering over dropdown items?
        if let Some(idx) = self.open_menu {
            let items = self.items_for(idx);
            let mut menu_x = rect.x + 6.0;
            for (i, label) in MENUS.iter().enumerate() {
                if i == idx { break; }
                menu_x += menu_label_w(label);
            }
            let dw = 200.0;
            let item_h = 28.0;
            let dy = rect.y + rect.h;

            if mx >= menu_x && mx < menu_x + dw {
                let rel_y = my - dy - 2.0;
                if rel_y >= 0.0 {
                    let item_idx = (rel_y / item_h) as usize;
                    if item_idx < items.len() {
                        self.hover_item = Some(item_idx);
                    }
                }
            }
        }
    }

    /// Render the menu bar background and labels (NOT the dropdown).
    pub fn render(
        &self,
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        rect: Rect,
        scale: f32,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(240, 240, 240, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        // Bottom border
        paint.set_color_rgba8(200, 200, 200, 255);
        fill(pix, &paint, rect.x, rect.y + rect.h - 1.0, rect.w, 1.0, s);

        // Menu labels
        let text_color = CosmicColor::rgba(50, 50, 50, 255);
        let mut x = rect.x + 6.0;
        for (i, label) in MENUS.iter().enumerate() {
            let lw = menu_label_w(label);

            // Highlight if open or hovered
            if self.open_menu == Some(i) || self.hover_menu == Some(i) {
                paint.set_color_rgba8(0, 102, 204, 50);
                fill(pix, &paint, x, rect.y, lw, rect.h, s);
            }

            draw_text(pix, fs, sc, label, x + 10.0, rect.y + 5.0, 14.0, text_color, s);
            x += lw;
        }
    }

    /// Render the dropdown overlay — call this LAST so it draws on top of everything.
    pub fn render_dropdown_overlay(
        &self,
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        rect: Rect,
        scale: f32,
    ) {
        if let Some(idx) = self.open_menu {
            self.render_dropdown(pix, fs, sc, idx, rect, scale);
        }
    }

    fn render_dropdown(
        &self,
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        menu_idx: usize,
        rect: Rect,
        s: f32,
    ) {
        let items = self.items_for(menu_idx);
        let mut menu_x = rect.x + 6.0;
        for (i, label) in MENUS.iter().enumerate() {
            if i == menu_idx { break; }
            menu_x += menu_label_w(label);
        }

        let dw = 200.0;
        let item_h = 28.0;
        let dh = items.len() as f32 * item_h + 4.0;
        let dy = rect.y + rect.h;

        let mut paint = Paint::default();

        // Shadow
        paint.set_color_rgba8(0, 0, 0, 30);
        fill(pix, &paint, menu_x + 2.0, dy + 2.0, dw, dh, s);

        // Background
        paint.set_color_rgba8(255, 255, 255, 255);
        fill(pix, &paint, menu_x, dy, dw, dh, s);

        // Border
        paint.set_color_rgba8(200, 200, 200, 255);
        fill(pix, &paint, menu_x, dy, dw, 1.0, s);

        // Items
        let text_color = CosmicColor::rgba(40, 40, 40, 255);
        for (i, (label, _)) in items.iter().enumerate() {
            let iy = dy + 2.0 + i as f32 * item_h;
            // Hover highlight
            if self.hover_item == Some(i) {
                paint.set_color_rgba8(0, 102, 204, 255);
                fill(pix, &paint, menu_x, iy, dw, item_h, s);
                let white = CosmicColor::rgba(255, 255, 255, 255);
                draw_text(pix, fs, sc, label, menu_x + 14.0, iy + 5.0, 13.0, white, s);
            } else {
                draw_text(pix, fs, sc, label, menu_x + 14.0, iy + 5.0, 13.0, text_color, s);
            }
        }
    }

    fn items_for(&self, idx: usize) -> &[(&str, MenuAction)] {
        match idx {
            0 => FILE_ITEMS,
            1 => EDIT_ITEMS,
            2 => PROJECT_ITEMS,
            3 => RUN_ITEMS,
            _ => &[],
        }
    }

    /// Handle click at logical (mx, my).
    pub fn handle_click(&mut self, mx: f32, my: f32, rect: Rect) -> Option<MenuAction> {
        // Menu bar labels
        if my >= rect.y && my < rect.y + rect.h {
            let mut x = rect.x + 6.0;
            for (i, label) in MENUS.iter().enumerate() {
                let lw = menu_label_w(label);
                if mx >= x && mx < x + lw {
                    self.open_menu = if self.open_menu == Some(i) { None } else { Some(i) };
                    return None;
                }
                x += lw;
            }
            self.open_menu = None;
            return None;
        }

        // Dropdown items
        if let Some(idx) = self.open_menu {
            let items = self.items_for(idx);
            let mut menu_x = rect.x + 6.0;
            for (i, label) in MENUS.iter().enumerate() {
                if i == idx { break; }
                menu_x += menu_label_w(label);
            }

            let dw = 200.0;
            let item_h = 28.0;
            let dy = rect.y + rect.h;

            if mx >= menu_x && mx < menu_x + dw {
                let rel_y = my - dy - 2.0;
                if rel_y >= 0.0 {
                    let item_idx = (rel_y / item_h) as usize;
                    if item_idx < items.len() {
                        let action = items[item_idx].1;
                        self.open_menu = None;
                        return Some(action);
                    }
                }
            }
            self.open_menu = None;
        }

        None
    }
}

/// Helper: fill a logical rect on the physical pixmap.
fn fill(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pix.fill_rect(r, paint, Transform::identity(), None);
    }
}
