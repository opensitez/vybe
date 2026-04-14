use vybe_widgets::{FontSystem, SwashCache, TextColor as CosmicColor};
use vybe_widgets::color_picker::{ColorPicker, ColorPickerEvent};
use tiny_skia::{Paint, Pixmap, Transform, Stroke, PathBuilder};
use uuid::Uuid;
use vybe_forms::{Form, Control, ControlType};

use crate::ide_text::{draw_text, draw_text_with_font, measure_text_with_font};


/// Rectangle in logical pixels.
#[derive(Clone, Copy, Debug)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}
impl Rect {
    pub fn contains(&self, px: f32, py: f32) -> bool {
        px >= self.x && px < self.x + self.w && py >= self.y && py < self.y + self.h
    }
}


pub struct FormLayout {
    pub toolbox: Rect,
    pub content: Rect,
    pub properties: Rect,
}

/// A property row or section header.
enum PropItem {
    Section(String),
    Row(String, String),
    CheckboxRow(String, bool),
    DropdownRow(String, String, Vec<String>),
}

#[derive(Clone, Copy, PartialEq)]
pub enum PropTab { Properties, Events }

const PROP_HEADER_H: f32 = 28.0;
const PROP_TAB_H: f32 = 26.0;
const PROP_ROW_H: f32 = 24.0;
const PROP_SECTION_H: f32 = 20.0;
const PROP_SCROLLBAR_W: f32 = 10.0;

#[derive(Debug, Clone, PartialEq)]
pub enum ControlTool {
    Pointer,
    Control(ControlType),
}

fn next_ctrl_name(ct: &ControlType, form: &Form) -> String {
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

fn snap(v: i32) -> i32 {
    (v / 10) * 10
}

const TITLE_H: f32 = 30.0;
const FORM_PADDING: f32 = 20.0;
const GRID_SIZE: f32 = 20.0;
const HANDLE_SZ: f32 = 6.0;
#[allow(dead_code)]
const MENU_BAR_H: f32 = 28.0;
#[allow(dead_code)]
const TOOLBAR_H: f32 = 36.0;

// ── Menu bar ──────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MenuAction {
    NewProject, OpenProject, SaveProject, SaveAs, Exit,
    Undo, Redo, Cut, Copy, Paste, Delete,
    AddForm, AddModule, AddExistingForm, AddExistingCode, AddResourceFile, ProjectProperties,
    RunProject, StopProject,
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
    ("Delete        Del", MenuAction::Delete),
    ("Cut           Ctrl+X", MenuAction::Cut),
    ("Copy          Ctrl+C", MenuAction::Copy),
    ("Paste         Ctrl+V", MenuAction::Paste),
];

const PROJECT_ITEMS: &[(&str, MenuAction)] = &[
    ("Add Form", MenuAction::AddForm),
    ("Add Module", MenuAction::AddModule),
    ("Add Existing Form...", MenuAction::AddExistingForm),
    ("Add Existing Code...", MenuAction::AddExistingCode),
    ("Add Resource File", MenuAction::AddResourceFile),
    ("Project Properties...", MenuAction::ProjectProperties),
];

const RUN_ITEMS: &[(&str, MenuAction)] = &[
    ("Start", MenuAction::RunProject),
    ("End", MenuAction::StopProject),
];

fn menu_label_w(label: &str) -> f32 {
    label.len() as f32 * 9.0 + 20.0
}

fn menu_items_for(idx: usize) -> &'static [(&'static str, MenuAction)] {
    match idx { 0 => FILE_ITEMS, 1 => EDIT_ITEMS, 2 => PROJECT_ITEMS, 3 => RUN_ITEMS, _ => &[] }
}

pub struct MenuBarState {
    pub open_menu: Option<usize>,
    pub hover_item: Option<usize>,
    pub hover_menu: Option<usize>,
}

impl MenuBarState {
    pub fn new() -> Self { Self { open_menu: None, hover_item: None, hover_menu: None } }

    pub fn handle_hover(&mut self, mx: f32, my: f32, rect: Rect) {
        self.hover_item = None;
        self.hover_menu = None;
        if my >= rect.y && my < rect.y + rect.h {
            let mut x = rect.x + 6.0;
            for (i, label) in MENUS.iter().enumerate() {
                let lw = menu_label_w(label);
                if mx >= x && mx < x + lw {
                    self.hover_menu = Some(i);
                    if self.open_menu.is_some() { self.open_menu = Some(i); }
                    return;
                }
                x += lw;
            }
            return;
        }
        if let Some(idx) = self.open_menu {
            let items = menu_items_for(idx);
            let mut menu_x = rect.x + 6.0;
            for (i, label) in MENUS.iter().enumerate() { if i == idx { break; } menu_x += menu_label_w(label); }
            let dw = 200.0; let item_h = 28.0; let dy = rect.y + rect.h;
            if mx >= menu_x && mx < menu_x + dw {
                let rel_y = my - dy - 2.0;
                if rel_y >= 0.0 {
                    let item_idx = (rel_y / item_h) as usize;
                    if item_idx < items.len() { self.hover_item = Some(item_idx); }
                }
            }
        }
    }

    pub fn handle_click(&mut self, mx: f32, my: f32, rect: Rect) -> Option<MenuAction> {
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
        if let Some(idx) = self.open_menu {
            let items = menu_items_for(idx);
            let mut menu_x = rect.x + 6.0;
            for (i, label) in MENUS.iter().enumerate() { if i == idx { break; } menu_x += menu_label_w(label); }
            let dw = 200.0; let item_h = 28.0; let dy = rect.y + rect.h;
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

    pub fn render(&self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: Rect, s: f32) {
        let mut paint = Paint::default();
        paint.set_color_rgba8(240, 240, 240, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);
        paint.set_color_rgba8(200, 200, 200, 255);
        fill(pix, &paint, rect.x, rect.y + rect.h - 1.0, rect.w, 1.0, s);
        let text_color = CosmicColor::rgba(50, 50, 50, 255);
        let mut x = rect.x + 6.0;
        for (i, label) in MENUS.iter().enumerate() {
            let lw = menu_label_w(label);
            if self.open_menu == Some(i) || self.hover_menu == Some(i) {
                paint.set_color_rgba8(0, 102, 204, 50);
                fill(pix, &paint, x, rect.y, lw, rect.h, s);
            }
            draw_text(pix, fs, sc, label, x + 10.0, rect.y + 5.0, 14.0, text_color, s);
            x += lw;
        }
    }

    pub fn render_dropdown_overlay(&self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: Rect, s: f32) {
        let idx = match self.open_menu { Some(i) => i, None => return };
        let items = menu_items_for(idx);
        let mut menu_x = rect.x + 6.0;
        for (i, label) in MENUS.iter().enumerate() { if i == idx { break; } menu_x += menu_label_w(label); }
        let dw = 200.0; let item_h = 28.0;
        let dh = items.len() as f32 * item_h + 4.0;
        let dy = rect.y + rect.h;
        let mut paint = Paint::default();
        paint.set_color_rgba8(0, 0, 0, 30);
        fill(pix, &paint, menu_x + 2.0, dy + 2.0, dw, dh, s);
        paint.set_color_rgba8(255, 255, 255, 255);
        fill(pix, &paint, menu_x, dy, dw, dh, s);
        paint.set_color_rgba8(200, 200, 200, 255);
        fill(pix, &paint, menu_x, dy, dw, 1.0, s);
        let text_color = CosmicColor::rgba(40, 40, 40, 255);
        for (i, (label, _)) in items.iter().enumerate() {
            let iy = dy + 2.0 + i as f32 * item_h;
            if self.hover_item == Some(i) {
                paint.set_color_rgba8(0, 102, 204, 255);
                fill(pix, &paint, menu_x, iy, dw, item_h, s);
                draw_text(pix, fs, sc, label, menu_x + 14.0, iy + 5.0, 13.0, CosmicColor::rgba(255, 255, 255, 255), s);
            } else {
                draw_text(pix, fs, sc, label, menu_x + 14.0, iy + 5.0, 13.0, text_color, s);
            }
        }
    }
}

// ── Toolbar ──────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ToolbarAction {
    Run, Stop, ViewDesigner, ViewCode, AddForm, AddCode, Save,
}

struct TButton { label: &'static str, action: Option<ToolbarAction>, is_sep: bool }

const TBUTTONS: &[TButton] = &[
    TButton { label: "\u{1F4BE} Save", action: Some(ToolbarAction::Save), is_sep: false },
    TButton { label: "",               action: None, is_sep: true },
    TButton { label: "\u{25B6} Start", action: Some(ToolbarAction::Run), is_sep: false },
    TButton { label: "\u{25A0} End",   action: Some(ToolbarAction::Stop), is_sep: false },
    TButton { label: "",               action: None, is_sep: true },
    TButton { label: "Designer",       action: Some(ToolbarAction::ViewDesigner), is_sep: false },
    TButton { label: "Code",           action: Some(ToolbarAction::ViewCode), is_sep: false },
    TButton { label: "",               action: None, is_sep: true },
    TButton { label: "+ Form",         action: Some(ToolbarAction::AddForm), is_sep: false },
    TButton { label: "+ Code",         action: Some(ToolbarAction::AddCode), is_sep: false },
];

pub fn render_toolbar_pub(pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: Rect, s: f32) {
    render_toolbar(pix, fs, sc, rect, s);
}

pub fn toolbar_handle_click_pub(mx: f32, my: f32, rect: Rect) -> Option<ToolbarAction> {
    toolbar_handle_click(mx, my, rect)
}

fn render_toolbar(pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: Rect, s: f32) {
    let mut paint = Paint::default();
    paint.set_color_rgba8(240, 240, 240, 255);
    fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);
    paint.set_color_rgba8(204, 204, 204, 255);
    fill(pix, &paint, rect.x, rect.y + rect.h - 1.0, rect.w, 1.0, s);
    let text_color = CosmicColor::rgba(50, 50, 50, 255);
    let mut x = rect.x + 8.0;
    let btn_h = 26.0;
    let btn_y = rect.y + (rect.h - btn_h) / 2.0;
    for btn in TBUTTONS {
        if btn.is_sep {
            paint.set_color_rgba8(200, 200, 200, 255);
            fill(pix, &paint, x, btn_y + 3.0, 1.0, btn_h - 6.0, s);
            x += 10.0;
            continue;
        }
        let btn_w = btn.label.len() as f32 * 8.0 + 20.0;
        paint.set_color_rgba8(255, 255, 255, 255);
        fill(pix, &paint, x, btn_y, btn_w, btn_h, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, x, btn_y, btn_w, 1.0, s);
        fill(pix, &paint, x, btn_y + btn_h - 1.0, btn_w, 1.0, s);
        fill(pix, &paint, x, btn_y, 1.0, btn_h, s);
        fill(pix, &paint, x + btn_w - 1.0, btn_y, 1.0, btn_h, s);
        draw_text(pix, fs, sc, btn.label, x + 10.0, btn_y + 5.0, 13.0, text_color, s);
        x += btn_w + 4.0;
    }
}

fn toolbar_handle_click(mx: f32, my: f32, rect: Rect) -> Option<ToolbarAction> {
    let btn_h = 26.0;
    let btn_y = rect.y + (rect.h - btn_h) / 2.0;
    if my < btn_y || my > btn_y + btn_h { return None; }
    let mut x = rect.x + 8.0;
    for btn in TBUTTONS {
        if btn.is_sep { x += 10.0; continue; }
        let btn_w = btn.label.len() as f32 * 8.0 + 20.0;
        if mx >= x && mx < x + btn_w { return btn.action; }
        x += btn_w + 4.0;
    }
    None
}

#[derive(Clone, Copy, PartialEq)]
pub enum ResizeHandle {
    TopLeft, Top, TopRight,
    Left, Right,
    BottomLeft, Bottom, BottomRight,
}

pub struct FormDesignerState {
    pub form: Form,
    pub selected_controls: Vec<Uuid>,
    pub drag_start: Option<(f32, f32)>,
    pub drag_offset: Option<(f32, f32)>,
    pub dragging: bool,
    pub resize_handle: Option<ResizeHandle>,
    pub resize_initial: Option<(i32, i32, i32, i32)>,
    pub drag_initial_bounds: Vec<(Uuid, i32, i32)>,
    pub lasso_start: Option<(f32, f32)>,
    pub lasso_current: Option<(f32, f32)>,
    pub scroll_x: f32,
    pub scroll_y: f32,
    pub toolbox: ToolboxState,
    pub prop_tab: PropTab,
    pub prop_scroll_y: f32,
    pub menu_bar: MenuBarState,
    pub color_picker: ColorPicker,
    pub color_picker_prop: Option<String>,
}

impl FormDesignerState {
    pub fn new() -> Self {
        let mut form = Form::new("Form1".to_string());
        form.width = 600;
        form.height = 400;
        Self {
            form,
            selected_controls: Vec::new(),
            drag_start: None,
            drag_offset: None,
            dragging: false,
            drag_initial_bounds: Vec::new(),
            resize_handle: None,
            resize_initial: None,
            lasso_start: None,
            lasso_current: None,
            scroll_x: 0.0,
            scroll_y: 0.0,
            toolbox: ToolboxState::new(),
            prop_tab: PropTab::Properties,
            prop_scroll_y: 0.0,
            menu_bar: MenuBarState::new(),
            color_picker: ColorPicker::new(),
            color_picker_prop: None,
        }
    }

    /// Centralized layout — all render and hit test code must use this.
    pub fn layout(&self, rect: Rect) -> FormLayout {
        let toolbox_w = 180.0;
        let properties_w = 220.0;
        let content_w = (rect.w - toolbox_w - properties_w).max(0.0);
        FormLayout {
            toolbox: Rect { x: rect.x, y: rect.y, w: toolbox_w, h: rect.h },
            content: Rect { x: rect.x + toolbox_w, y: rect.y, w: content_w, h: rect.h },
            properties: Rect { x: rect.x + rect.w - properties_w, y: rect.y, w: properties_w, h: rect.h },
        }
    }

    fn form_client_origin(&self, content_rect: Rect) -> (f32, f32) {
        (
            content_rect.x + FORM_PADDING - self.scroll_x,
            content_rect.y + FORM_PADDING - self.scroll_y + TITLE_H,
        )
    }

    pub fn render(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        rect: Rect, scale: f32,
    ) {
        let lay = self.layout(rect);
        let content_rect = lay.content;
        let toolbox_rect = lay.toolbox;
        let properties_rect = lay.properties;

        let s = scale;
        let mut paint = Paint::default();

        // Workspace background
        paint.set_color_rgba8(45, 45, 48, 255);
        fill(pix, &paint, content_rect.x, content_rect.y, content_rect.w, content_rect.h, s);

        let form_x = content_rect.x + FORM_PADDING - self.scroll_x;
        let form_y = content_rect.y + FORM_PADDING - self.scroll_y;
        let form_w = self.form.width as f32;
        let form_h = self.form.height as f32;

        // Shadow
        paint.set_color_rgba8(0, 0, 0, 40);
        fill(pix, &paint, form_x + 3.0, form_y + 3.0, form_w, form_h, s);

        // Title bar (blue gradient-like)
        paint.set_color_rgba8(0, 120, 212, 255);
        fill(pix, &paint, form_x, form_y, form_w, TITLE_H, s);

        let title_color = CosmicColor::rgba(255, 255, 255, 255);
        draw_text(pix, fs, sc, &self.form.text, form_x + 10.0, form_y + 7.0, 13.0, title_color, s);

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
        if let Some(ref hex) = self.form.back_color {
            if let Some(c) = vybe_widgets::color_picker::PickedColor::from_hex(hex) {
                paint.set_color_rgba8(c.r, c.g, c.b, c.a);
            } else {
                paint.set_color_rgba8(240, 240, 240, 255);
            }
        } else {
            paint.set_color_rgba8(240, 240, 240, 255);
        }
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

        // Controls
        let (cx0, cy0) = self.form_client_origin(content_rect);
        self.render_controls_recursive(pix, fs, sc, None, cx0, cy0, s, 0);

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

        // Component tray
        self.render_tray(pix, fs, sc, form_x, form_y + form_h, form_w, s);

        // Toolbox (left panel)
        self.toolbox.render(pix, fs, sc, toolbox_rect, scale);

        // Properties panel (right panel)
        self.render_properties(pix, fs, sc, properties_rect, scale);
    }

    const TRAY_H: f32 = 48.0;
    const TRAY_ITEM_W: f32 = 90.0;
    const TRAY_ITEM_H: f32 = 36.0;
    const TRAY_GAP: f32 = 6.0;

    fn non_visual_controls(&self) -> Vec<&Control> {
        self.form.controls.iter().filter(|c| c.control_type.is_non_visual()).collect()
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
        tray_x: f32, tray_top: f32, tray_w: f32, s: f32,
    ) {
        let non_visuals = self.non_visual_controls();
        if non_visuals.is_empty() { return; }

        let mut paint = Paint::default();
        let tray_y = tray_top + 8.0;

        paint.set_color_rgba8(245, 245, 245, 255);
        fill(pix, &paint, tray_x, tray_y, tray_w, Self::TRAY_H, s);
        paint.set_color_rgba8(200, 200, 200, 255);
        fill(pix, &paint, tray_x, tray_y, tray_w, 1.0, s);

        let dim = CosmicColor::rgba(120, 120, 120, 255);
        draw_text(pix, fs, sc, "Components", tray_x + 4.0, tray_y + 2.0, 9.0, dim, s);

        let text_color = CosmicColor::rgba(50, 50, 50, 255);
        let mut ix = tray_x + 4.0;
        let iy = tray_y + 12.0;

        for ctrl in &non_visuals {
            let is_sel = self.selected_controls.contains(&ctrl.id);

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

            let icon = Self::tray_icon(&ctrl.control_type);
            draw_text(pix, fs, sc, icon, ix + 4.0, iy + 4.0, 14.0, text_color, s);

            let name = if ctrl.name.len() > 10 {
                format!("{}...", &ctrl.name[..8])
            } else {
                ctrl.name.clone()
            };
            draw_text(pix, fs, sc, &name, ix + 4.0, iy + 20.0, 9.0, text_color, s);

            ix += Self::TRAY_ITEM_W + Self::TRAY_GAP;
        }
    }

    fn prop(key: &str, val: &str) -> PropItem {
        PropItem::Row(key.into(), val.into())
    }

    fn get_prop(ctrl: &Control, key: &str, default: &str) -> String {
        ctrl.properties.get_string(key).unwrap_or(default).to_string()
    }

    fn collect_props(&self) -> Vec<PropItem> {
        use ControlType::*;
        if let Some(id) = self.selected_controls.first() {
            if let Some(ctrl) = self.form.controls.iter().find(|c| c.id == *id) {
                let mut items = vec![
                    PropItem::Section("Basic".into()),
                    Self::prop("Name", &ctrl.name),
                    Self::prop("Type", &format!("{:?}", ctrl.control_type)),
                ];

                if !ctrl.control_type.is_non_visual() {
                    items.push(Self::prop("Left", &format!("{}", ctrl.bounds.x)));
                    items.push(Self::prop("Top", &format!("{}", ctrl.bounds.y)));
                    items.push(Self::prop("Width", &format!("{}", ctrl.bounds.width)));
                    items.push(Self::prop("Height", &format!("{}", ctrl.bounds.height)));
                }

                // Appearance
                items.push(PropItem::Section("Appearance".into()));
                items.push(Self::prop("Text", &Self::get_prop(ctrl, "Text", "")));
                items.push(Self::prop("BackColor", &Self::get_prop(ctrl, "BackColor", "")));
                items.push(Self::prop("ForeColor", &Self::get_prop(ctrl, "ForeColor", "")));
                items.push(Self::prop("Font", &Self::get_prop(ctrl, "Font", "")));
                items.push(Self::prop("TabIndex", &format!("{}", ctrl.tab_index)));
                items.push(PropItem::CheckboxRow("Enabled".into(),
                    Self::get_prop(ctrl, "Enabled", "True").eq_ignore_ascii_case("true")));
                items.push(PropItem::CheckboxRow("Visible".into(),
                    Self::get_prop(ctrl, "Visible", "True").eq_ignore_ascii_case("true")));

                // Control-specific
                match ctrl.control_type {
                    CheckBox | RadioButton => {
                        items.push(PropItem::Section("State".into()));
                        items.push(PropItem::CheckboxRow("Checked".into(),
                            Self::get_prop(ctrl, "Checked", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::DropdownRow("CheckState".into(),
                            Self::get_prop(ctrl, "CheckState", "Unchecked"),
                            vec!["Unchecked".into(), "Checked".into(), "Indeterminate".into()]));
                    }
                    TextBox => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::CheckboxRow("Multiline".into(),
                            Self::get_prop(ctrl, "Multiline", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("ReadOnly".into(),
                            Self::get_prop(ctrl, "ReadOnly", "False").eq_ignore_ascii_case("true")));
                        items.push(Self::prop("MaxLength", &Self::get_prop(ctrl, "MaxLength", "0")));
                        items.push(Self::prop("PasswordChar", &Self::get_prop(ctrl, "PasswordChar", "")));
                    }
                    RichTextBox => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::CheckboxRow("ReadOnly".into(),
                            Self::get_prop(ctrl, "ReadOnly", "False").eq_ignore_ascii_case("true")));
                    }
                    NumericUpDown => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("Minimum", &Self::get_prop(ctrl, "Minimum", "0")));
                        items.push(Self::prop("Maximum", &Self::get_prop(ctrl, "Maximum", "100")));
                        items.push(Self::prop("Increment", &Self::get_prop(ctrl, "Increment", "1")));
                        items.push(Self::prop("DecimalPlaces", &Self::get_prop(ctrl, "DecimalPlaces", "0")));
                    }
                    TrackBar => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("Minimum", &Self::get_prop(ctrl, "Minimum", "0")));
                        items.push(Self::prop("Maximum", &Self::get_prop(ctrl, "Maximum", "10")));
                        items.push(Self::prop("SmallChange", &Self::get_prop(ctrl, "SmallChange", "1")));
                        items.push(Self::prop("LargeChange", &Self::get_prop(ctrl, "LargeChange", "5")));
                    }
                    ProgressBar => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("Minimum", &Self::get_prop(ctrl, "Minimum", "0")));
                        items.push(Self::prop("Maximum", &Self::get_prop(ctrl, "Maximum", "100")));
                    }
                    MaskedTextBox => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("Mask", &Self::get_prop(ctrl, "Mask", "")));
                        items.push(Self::prop("PromptChar", &Self::get_prop(ctrl, "PromptChar", "_")));
                    }
                    ComboBox | ListBox => {
                        items.push(PropItem::Section("Behavior".into()));
                        if ctrl.control_type == ComboBox {
                            items.push(PropItem::DropdownRow("DropDownStyle".into(),
                                Self::get_prop(ctrl, "DropDownStyle", "DropDown"),
                                vec!["Simple".into(), "DropDown".into(), "DropDownList".into()]));
                        }
                        if ctrl.control_type == ListBox {
                            items.push(PropItem::CheckboxRow("Sorted".into(),
                                Self::get_prop(ctrl, "Sorted", "False").eq_ignore_ascii_case("true")));
                            items.push(PropItem::DropdownRow("SelectionMode".into(),
                                Self::get_prop(ctrl, "SelectionMode", "One"),
                                vec!["None".into(), "One".into(), "MultiSimple".into(), "MultiExtended".into()]));
                        }
                    }
                    TreeView => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::CheckboxRow("CheckBoxes".into(),
                            Self::get_prop(ctrl, "CheckBoxes", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("ShowLines".into(),
                            Self::get_prop(ctrl, "ShowLines", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("ShowRootLines".into(),
                            Self::get_prop(ctrl, "ShowRootLines", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("ShowPlusMinus".into(),
                            Self::get_prop(ctrl, "ShowPlusMinus", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("LabelEdit".into(),
                            Self::get_prop(ctrl, "LabelEdit", "False").eq_ignore_ascii_case("true")));
                    }
                    ListView => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::DropdownRow("View".into(),
                            Self::get_prop(ctrl, "View", "Details"),
                            vec!["LargeIcon".into(), "Details".into(), "SmallIcon".into(), "List".into(), "Tile".into()]));
                        items.push(PropItem::CheckboxRow("FullRowSelect".into(),
                            Self::get_prop(ctrl, "FullRowSelect", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("GridLines".into(),
                            Self::get_prop(ctrl, "GridLines", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("CheckBoxes".into(),
                            Self::get_prop(ctrl, "CheckBoxes", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("MultiSelect".into(),
                            Self::get_prop(ctrl, "MultiSelect", "True").eq_ignore_ascii_case("true")));
                    }
                    DataGridView => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::CheckboxRow("ReadOnly".into(),
                            Self::get_prop(ctrl, "ReadOnly", "False").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("AllowUserToAddRows".into(),
                            Self::get_prop(ctrl, "AllowUserToAddRows", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("AllowUserToDeleteRows".into(),
                            Self::get_prop(ctrl, "AllowUserToDeleteRows", "True").eq_ignore_ascii_case("true")));
                        items.push(PropItem::CheckboxRow("AutoGenerateColumns".into(),
                            Self::get_prop(ctrl, "AutoGenerateColumns", "True").eq_ignore_ascii_case("true")));
                    }
                    TabControl => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::DropdownRow("Alignment".into(),
                            Self::get_prop(ctrl, "Alignment", "Top"),
                            vec!["Top".into(), "Bottom".into(), "Left".into(), "Right".into()]));
                        items.push(PropItem::CheckboxRow("Multiline".into(),
                            Self::get_prop(ctrl, "Multiline", "False").eq_ignore_ascii_case("true")));
                    }
                    DateTimePicker => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(PropItem::DropdownRow("Format".into(),
                            Self::get_prop(ctrl, "Format", "Long"),
                            vec!["Long".into(), "Short".into(), "Time".into(), "Custom".into()]));
                    }
                    WebBrowser => {
                        items.push(PropItem::Section("Behavior".into()));
                        items.push(Self::prop("URL", &Self::get_prop(ctrl, "URL", "")));
                    }
                    _ => {}
                }

                // Data Binding
                items.push(PropItem::Section("Data".into()));
                let is_non_visual = ctrl.control_type.is_non_visual();
                let has_complex = matches!(ctrl.control_type,
                    DataGridView | ListBox | ComboBox | BindingNavigator |
                    BindingSourceComponent);

                let mut bs_options = vec!["(none)".to_string()];
                let mut ds_options = vec!["(none)".to_string()];
                for c in &self.form.controls {
                    if matches!(c.control_type, BindingSourceComponent) && c.id != ctrl.id {
                        bs_options.push(c.name.clone());
                        ds_options.push(c.name.clone());
                    }
                    if matches!(c.control_type, DataAdapterComponent | DataSetComponent | DataTableComponent) {
                        ds_options.push(c.name.clone());
                    }
                }

                if !is_non_visual && !has_complex {
                    let bindable = match ctrl.control_type {
                        CheckBox | RadioButton => "Checked",
                        PictureBox => "ImageLocation",
                        _ => "Text",
                    };
                    items.push(PropItem::DropdownRow("DataBindings.Source".into(),
                        Self::get_prop(ctrl, "DataBindings.Source", ""), bs_options.clone()));
                    items.push(Self::prop(&format!("Bind: {}", bindable),
                        &Self::get_prop(ctrl, &format!("DataBindings.{}", bindable), "")));
                }

                if has_complex && ctrl.control_type != BindingNavigator {
                    items.push(PropItem::DropdownRow("DataSource".into(),
                        Self::get_prop(ctrl, "DataSource", ""), ds_options.clone()));
                    items.push(Self::prop("DataMember", &Self::get_prop(ctrl, "DataMember", "")));
                }

                if matches!(ctrl.control_type, ListBox | ComboBox) {
                    items.push(Self::prop("DisplayMember", &Self::get_prop(ctrl, "DisplayMember", "")));
                    items.push(Self::prop("ValueMember", &Self::get_prop(ctrl, "ValueMember", "")));
                }

                if ctrl.control_type == BindingNavigator {
                    items.push(PropItem::DropdownRow("BindingSource".into(),
                        Self::get_prop(ctrl, "BindingSource", ""), bs_options));
                }

                if ctrl.control_type == BindingSourceComponent {
                    items.push(Self::prop("DataSource", &Self::get_prop(ctrl, "DataSource", "")));
                    items.push(Self::prop("DataMember", &Self::get_prop(ctrl, "DataMember", "")));
                    items.push(Self::prop("Filter", &Self::get_prop(ctrl, "Filter", "")));
                    items.push(Self::prop("Sort", &Self::get_prop(ctrl, "Sort", "")));
                }

                if ctrl.control_type == DataSetComponent {
                    items.push(Self::prop("DataSetName", &Self::get_prop(ctrl, "DataSetName", "NewDataSet")));
                }

                if ctrl.control_type == DataTableComponent {
                    items.push(Self::prop("TableName", &Self::get_prop(ctrl, "TableName", "Table1")));
                }

                if ctrl.control_type == DataAdapterComponent {
                    items.push(Self::prop("SelectCommand", &Self::get_prop(ctrl, "SelectCommand", "")));
                    items.push(Self::prop("ConnectionString", &Self::get_prop(ctrl, "ConnectionString", "")));
                    items.push(PropItem::Section("Connection Builder".into()));
                    items.push(PropItem::DropdownRow("DbType".into(),
                        Self::get_prop(ctrl, "DbType", "SQLite"),
                        vec!["SQLite".into(), "PostgreSQL".into(), "MySQL".into()]));
                    let db_type = Self::get_prop(ctrl, "DbType", "SQLite");
                    if db_type == "SQLite" {
                        items.push(Self::prop("DbPath", &Self::get_prop(ctrl, "DbPath", "")));
                    }
                }

                return items;
            }
        }
        // No control selected — show form properties
        vec![
            PropItem::Section("Form".into()),
            Self::prop("Name", &self.form.name),
            Self::prop("Text", &self.form.text),
            Self::prop("Width", &format!("{}", self.form.width)),
            Self::prop("Height", &format!("{}", self.form.height)),
            PropItem::Section("Appearance".into()),
            Self::prop("BackColor", &self.form.back_color.clone().unwrap_or_default()),
            Self::prop("ForeColor", &self.form.fore_color.clone().unwrap_or_default()),
            Self::prop("Font", &self.form.font.clone().unwrap_or_default()),
        ]
    }

    fn collect_events(&self) -> Vec<PropItem> {
        use ControlType::*;
        let ct = if let Some(id) = self.selected_controls.first() {
            self.form.controls.iter().find(|c| c.id == *id).map(|c| c.control_type.clone())
        } else {
            None // form-level events
        };

        let events: &[&str] = match ct.as_ref() {
            Some(Button) => &["Click", "MouseDown", "MouseUp", "MouseMove", "MouseEnter", "MouseLeave", "GotFocus", "LostFocus", "KeyDown", "KeyUp", "KeyPress", "EnabledChanged", "VisibleChanged", "Paint"],
            Some(TextBox) | Some(MaskedTextBox) | Some(RichTextBox) => &["TextChanged", "KeyPress", "KeyDown", "KeyUp", "GotFocus", "LostFocus", "Click", "Enter", "Leave", "Validating", "Validated"],
            Some(Label) | Some(LinkLabel) => &["Click", "DoubleClick", "MouseEnter", "MouseLeave"],
            Some(CheckBox) | Some(RadioButton) => &["CheckedChanged", "Click", "GotFocus", "LostFocus", "KeyPress", "EnabledChanged"],
            Some(ListBox) => &["SelectedIndexChanged", "Click", "DoubleClick", "GotFocus", "LostFocus", "KeyPress", "KeyDown"],
            Some(ComboBox) => &["SelectedIndexChanged", "SelectedValueChanged", "TextChanged", "DropDown", "DropDownClosed", "Click", "GotFocus", "LostFocus", "KeyPress", "KeyDown"],
            Some(TreeView) => &["AfterSelect", "BeforeSelect", "AfterExpand", "AfterCollapse", "NodeMouseClick", "NodeMouseDoubleClick", "AfterCheck", "Click", "DoubleClick", "KeyDown"],
            Some(ListView) => &["SelectedIndexChanged", "ItemActivate", "ColumnClick", "ItemCheck", "Click", "DoubleClick", "KeyDown"],
            Some(DataGridView) => &["CellClick", "CellDoubleClick", "CellValueChanged", "CellEndEdit", "CellBeginEdit", "SelectionChanged", "RowEnter", "RowLeave", "DataBindingComplete", "KeyDown"],
            Some(TabControl) => &["SelectedIndexChanged", "Selected", "Deselecting", "Click", "DoubleClick"],
            Some(Panel) | Some(Frame) => &["Click", "DoubleClick", "MouseDown", "MouseUp", "MouseMove", "MouseEnter", "MouseLeave", "Paint", "Resize"],
            Some(PictureBox) => &["Click", "DoubleClick", "MouseDown", "MouseUp", "MouseMove", "Paint"],
            Some(ProgressBar) => &["Click", "ValueChanged"],
            Some(NumericUpDown) => &["ValueChanged", "KeyPress", "KeyDown", "GotFocus", "LostFocus", "Validating"],
            Some(DateTimePicker) => &["ValueChanged", "DropDown", "DropDownClosed", "GotFocus", "LostFocus", "KeyPress"],
            Some(TrackBar) => &["Scroll", "ValueChanged", "MouseDown", "MouseUp", "GotFocus", "LostFocus"],
            Some(MenuStrip) | Some(ContextMenuStrip) => &["ItemClicked", "Click"],
            Some(StatusStrip) => &["ItemClicked", "Click"],
            Some(ToolStrip) => &["ItemClicked", "ButtonClick", "Click"],
            Some(WebBrowser) => &["DocumentCompleted", "Navigating", "Navigated", "ProgressChanged", "Click"],
            Some(SplitContainer) => &["SplitterMoved", "SplitterMoving", "Click", "DoubleClick"],
            Some(MonthCalendar) => &["DateChanged", "DateSelected", "Click", "DoubleClick", "GotFocus", "LostFocus"],
            Some(BindingSourceComponent) => &["CurrentChanged", "PositionChanged", "DataSourceChanged"],
            Some(BindingNavigator) => &["Click"],
            Some(Timer) => &["Tick"],
            None => &["Load", "Shown", "Activated", "Deactivate", "FormClosing", "FormClosed", "Resize", "Paint", "Click", "DoubleClick", "KeyDown", "KeyUp", "KeyPress", "MouseClick", "MouseDown", "MouseUp"],
            _ => &["Click", "DoubleClick", "MouseDown", "MouseUp", "GotFocus", "LostFocus"],
        };

        events.iter().map(|&ev| PropItem::Row(ev.into(), String::new())).collect()
    }

    fn items_height(items: &[PropItem]) -> f32 {
        items.iter().map(|i| match i {
            PropItem::Section(_) => PROP_SECTION_H,
            PropItem::Row(_, _) | PropItem::CheckboxRow(_, _) | PropItem::DropdownRow(_, _, _) => PROP_ROW_H,
        }).sum()
    }

    fn render_properties(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        rect: Rect, scale: f32,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        // Left border
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y, 1.0, rect.h, s);

        // Title
        let title_color = CosmicColor::rgba(50, 50, 50, 255);
        draw_text(pix, fs, sc, "Properties", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + PROP_HEADER_H - 1.0, rect.w, 1.0, s);

        // Tabs
        let tab_y = rect.y + PROP_HEADER_H;
        let tab_w = (rect.w - 2.0) / 2.0;
        let text_color = CosmicColor::rgba(30, 30, 30, 255);
        let dim_color = CosmicColor::rgba(100, 100, 100, 255);

        if self.prop_tab == PropTab::Properties {
            paint.set_color_rgba8(227, 242, 253, 255);
        } else {
            paint.set_color_rgba8(245, 245, 245, 255);
        }
        fill(pix, &paint, rect.x + 1.0, tab_y, tab_w, PROP_TAB_H, s);
        draw_text(pix, fs, sc, "Properties", rect.x + 14.0, tab_y + 5.0, 12.0,
            if self.prop_tab == PropTab::Properties { text_color } else { dim_color }, s);

        if self.prop_tab == PropTab::Events {
            paint.set_color_rgba8(227, 242, 253, 255);
        } else {
            paint.set_color_rgba8(245, 245, 245, 255);
        }
        fill(pix, &paint, rect.x + 1.0 + tab_w, tab_y, tab_w, PROP_TAB_H, s);
        draw_text(pix, fs, sc, "Events", rect.x + tab_w + 14.0, tab_y + 5.0, 12.0,
            if self.prop_tab == PropTab::Events { text_color } else { dim_color }, s);

        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, tab_y + PROP_TAB_H - 1.0, rect.w, 1.0, s);

        // Content area
        let content_top = tab_y + PROP_TAB_H;
        let content_h = rect.h - PROP_HEADER_H - PROP_TAB_H;
        let items = match self.prop_tab {
            PropTab::Properties => self.collect_props(),
            PropTab::Events => self.collect_events(),
        };

        let label_color = CosmicColor::rgba(80, 80, 80, 255);
        let val_color = CosmicColor::rgba(30, 30, 30, 255);
        let section_color = CosmicColor::rgba(100, 100, 100, 255);
        let val_x = rect.x + rect.w * 0.42;
        let mut y = content_top + 2.0 - self.prop_scroll_y;
        let mut entry_idx = 0usize;

        for item in &items {
            match item {
                PropItem::Section(label) => {
                    if y + PROP_SECTION_H > content_top && y < content_top + content_h {
                        paint.set_color_rgba8(240, 240, 240, 255);
                        fill(pix, &paint, rect.x + 1.0, y, rect.w - 2.0, PROP_SECTION_H, s);
                        draw_text(pix, fs, sc, label, rect.x + 8.0, y + 3.0, 10.0, section_color, s);
                    }
                    y += PROP_SECTION_H;
                }
                PropItem::Row(key, value) => {
                    if y + PROP_ROW_H > content_top && y < content_top + content_h {
                        if entry_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(pix, &paint, rect.x + 1.0, y, rect.w - 2.0, PROP_ROW_H, s);
                        }
                        draw_text(pix, fs, sc, key, rect.x + 8.0, y + 4.0, 11.0, label_color, s);
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(pix, &paint, val_x - 2.0, y, 1.0, PROP_ROW_H, s);
                        draw_text(pix, fs, sc, value, val_x + 4.0, y + 4.0, 11.0, val_color, s);

                        // Color swatch for BackColor / ForeColor
                        if key == "BackColor" || key == "ForeColor" {
                            let swatch_x = rect.x + rect.w - PROP_SCROLLBAR_W - 22.0;
                            let swatch_y = y + 3.0;
                            let swatch_sz = PROP_ROW_H - 6.0;
                            if let Some(c) = vybe_widgets::color_picker::PickedColor::from_hex(value) {
                                let mut sp = Paint::default();
                                sp.set_color_rgba8(c.r, c.g, c.b, c.a);
                                fill(pix, &sp, swatch_x, swatch_y, swatch_sz, swatch_sz, s);
                                sp.set_color_rgba8(160, 160, 160, 255);
                                stroke_rect(pix, &sp, swatch_x, swatch_y, swatch_sz, swatch_sz, s);
                            }
                        }

                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(pix, &paint, rect.x + 1.0, y + PROP_ROW_H - 1.0, rect.w - 2.0, 1.0, s);
                    }
                    entry_idx += 1;
                    y += PROP_ROW_H;
                }
                PropItem::CheckboxRow(key, checked) => {
                    if y + PROP_ROW_H > content_top && y < content_top + content_h {
                        if entry_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(pix, &paint, rect.x + 1.0, y, rect.w - 2.0, PROP_ROW_H, s);
                        }
                        draw_text(pix, fs, sc, key, rect.x + 8.0, y + 4.0, 11.0, label_color, s);
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(pix, &paint, val_x - 2.0, y, 1.0, PROP_ROW_H, s);
                        let cb_x = val_x + 4.0;
                        let cb_y = y + 4.0;
                        let cb_sz = 14.0;
                        paint.set_color_rgba8(255, 255, 255, 255);
                        fill(pix, &paint, cb_x, cb_y, cb_sz, cb_sz, s);
                        paint.set_color_rgba8(160, 160, 160, 255);
                        stroke_rect(pix, &paint, cb_x, cb_y, cb_sz, cb_sz, s);
                        if *checked {
                            paint.set_color_rgba8(0, 120, 212, 255);
                            fill(pix, &paint, cb_x + 3.0, cb_y + 3.0, cb_sz - 6.0, cb_sz - 6.0, s);
                        }
                        let label_text = if *checked { "True" } else { "False" };
                        draw_text(pix, fs, sc, label_text, cb_x + cb_sz + 4.0, y + 4.0, 11.0, val_color, s);
                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(pix, &paint, rect.x + 1.0, y + PROP_ROW_H - 1.0, rect.w - 2.0, 1.0, s);
                    }
                    entry_idx += 1;
                    y += PROP_ROW_H;
                }
                PropItem::DropdownRow(key, current, _options) => {
                    if y + PROP_ROW_H > content_top && y < content_top + content_h {
                        if entry_idx % 2 == 1 {
                            paint.set_color_rgba8(247, 247, 247, 255);
                            fill(pix, &paint, rect.x + 1.0, y, rect.w - 2.0, PROP_ROW_H, s);
                        }
                        draw_text(pix, fs, sc, key, rect.x + 8.0, y + 4.0, 11.0, label_color, s);
                        paint.set_color_rgba8(230, 230, 230, 255);
                        fill(pix, &paint, val_x - 2.0, y, 1.0, PROP_ROW_H, s);
                        let dd_w = rect.w - (val_x - rect.x) - PROP_SCROLLBAR_W - 2.0;
                        paint.set_color_rgba8(255, 255, 255, 255);
                        fill(pix, &paint, val_x, y + 1.0, dd_w, PROP_ROW_H - 2.0, s);
                        paint.set_color_rgba8(180, 180, 180, 255);
                        stroke_rect(pix, &paint, val_x, y + 1.0, dd_w, PROP_ROW_H - 2.0, s);
                        draw_text(pix, fs, sc, current, val_x + 4.0, y + 4.0, 11.0, val_color, s);
                        // Dropdown arrow
                        let arrow_x = val_x + dd_w - 14.0;
                        let arrow_y = y + PROP_ROW_H / 2.0 - 2.0;
                        paint.set_color_rgba8(80, 80, 80, 255);
                        fill(pix, &paint, arrow_x, arrow_y, 8.0, 1.0, s);
                        fill(pix, &paint, arrow_x + 1.0, arrow_y + 1.0, 6.0, 1.0, s);
                        fill(pix, &paint, arrow_x + 2.0, arrow_y + 2.0, 4.0, 1.0, s);
                        fill(pix, &paint, arrow_x + 3.0, arrow_y + 3.0, 2.0, 1.0, s);
                        paint.set_color_rgba8(235, 235, 235, 255);
                        fill(pix, &paint, rect.x + 1.0, y + PROP_ROW_H - 1.0, rect.w - 2.0, 1.0, s);
                    }
                    entry_idx += 1;
                    y += PROP_ROW_H;
                }
            }
        }

        if items.is_empty() {
            draw_text(pix, fs, sc, "No selection", rect.x + 10.0, content_top + 8.0, 12.0, dim_color, s);
        }

        // Overdraw header area to clip scrolled items
        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, PROP_HEADER_H + PROP_TAB_H, s);
        // Re-render header
        draw_text(pix, fs, sc, "Properties", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + PROP_HEADER_H - 1.0, rect.w, 1.0, s);
        // Re-render tabs
        if self.prop_tab == PropTab::Properties {
            paint.set_color_rgba8(227, 242, 253, 255);
        } else {
            paint.set_color_rgba8(245, 245, 245, 255);
        }
        fill(pix, &paint, rect.x + 1.0, tab_y, tab_w, PROP_TAB_H, s);
        draw_text(pix, fs, sc, "Properties", rect.x + 14.0, tab_y + 5.0, 12.0,
            if self.prop_tab == PropTab::Properties { text_color } else { dim_color }, s);
        if self.prop_tab == PropTab::Events {
            paint.set_color_rgba8(227, 242, 253, 255);
        } else {
            paint.set_color_rgba8(245, 245, 245, 255);
        }
        fill(pix, &paint, rect.x + 1.0 + tab_w, tab_y, tab_w, PROP_TAB_H, s);
        draw_text(pix, fs, sc, "Events", rect.x + tab_w + 14.0, tab_y + 5.0, 12.0,
            if self.prop_tab == PropTab::Events { text_color } else { dim_color }, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, tab_y + PROP_TAB_H - 1.0, rect.w, 1.0, s);

        // Scrollbar
        let total_h = Self::items_height(&items);
        let max_scroll = (total_h - content_h).max(0.0);
        if max_scroll > 0.0 {
            let sb_x = rect.x + rect.w - PROP_SCROLLBAR_W;
            paint.set_color_rgba8(235, 235, 235, 255);
            fill(pix, &paint, sb_x, content_top, PROP_SCROLLBAR_W, content_h, s);
            let visible_frac = (content_h / total_h).min(1.0);
            let thumb_h = (content_h * visible_frac).max(20.0);
            let scroll_frac = if max_scroll > 0.0 { self.prop_scroll_y / max_scroll } else { 0.0 };
            let thumb_y = content_top + scroll_frac * (content_h - thumb_h);
            paint.set_color_rgba8(190, 190, 190, 255);
            fill(pix, &paint, sb_x + 2.0, thumb_y, PROP_SCROLLBAR_W - 4.0, thumb_h, s);
        }

        // Left border (re-draw on top)
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y, 1.0, rect.h, s);

        // ── Color picker popup overlay ──
        if self.color_picker.open {
            let popup_x = rect.x + 10.0;
            let popup_y = rect.y + PROP_HEADER_H + PROP_TAB_H + 40.0;
            self.color_picker.render_popup(pix, popup_x, popup_y, s);
        }
    }

    pub fn handle_properties_click(&mut self, mx: f32, my: f32, rect: Rect) -> bool {
        if !rect.contains(mx, my) { return false; }

        // ── Color picker popup is open: route click to it first ──
        if self.color_picker.open {
            let popup_x = rect.x + 10.0;
            let popup_y = rect.y + PROP_HEADER_H + PROP_TAB_H + 40.0;
            match self.color_picker.handle_click(mx, my, popup_x, popup_y) {
                ColorPickerEvent::Changed(c) => {
                    let hex = c.to_hex();
                    if let Some(prop_name) = self.color_picker_prop.clone() {
                        self.apply_color_prop(&prop_name, &hex);
                    }
                    return true;
                }
                ColorPickerEvent::Closed => {
                    let hex = self.color_picker.color.to_hex();
                    if let Some(prop_name) = self.color_picker_prop.take() {
                        self.apply_color_prop(&prop_name, &hex);
                    }
                    return true;
                }
                ColorPickerEvent::None => { return true; }
            }
        }

        // Tab click
        let tab_y = rect.y + PROP_HEADER_H;
        if my >= tab_y && my < tab_y + PROP_TAB_H {
            let tab_w = (rect.w - 2.0) / 2.0;
            if mx < rect.x + 1.0 + tab_w {
                self.prop_tab = PropTab::Properties;
            } else {
                self.prop_tab = PropTab::Events;
            }
            self.prop_scroll_y = 0.0;
            return true;
        }

        // Property row click (value column)
        if self.prop_tab == PropTab::Properties {
            let content_top = tab_y + PROP_TAB_H;
            let val_x = rect.x + rect.w * 0.42;
            if my >= content_top && mx >= val_x {
                let items = self.collect_props();
                let mut y = content_top + 2.0 - self.prop_scroll_y;
                for item in &items {
                    match item {
                        PropItem::Section(_) => { y += PROP_SECTION_H; }
                        PropItem::Row(key, value) => {
                            if my >= y && my < y + PROP_ROW_H {
                                if key == "BackColor" || key == "ForeColor" {
                                    self.color_picker.set_from_hex(value);
                                    self.color_picker.open = true;
                                    self.color_picker_prop = Some(key.clone());
                                    return true;
                                }
                            }
                            y += PROP_ROW_H;
                        }
                        PropItem::CheckboxRow(_, _) | PropItem::DropdownRow(_, _, _) => { y += PROP_ROW_H; }
                    }
                }
            }
        }

        false
    }

    fn apply_color_prop(&mut self, prop_name: &str, hex: &str) {
        if self.selected_controls.is_empty() {
            // Apply to form itself
            match prop_name {
                "BackColor" => { self.form.back_color = Some(hex.to_string()); }
                "ForeColor" => { self.form.fore_color = Some(hex.to_string()); }
                _ => {}
            }
        } else {
            // Apply to selected control(s)
            for id in &self.selected_controls {
                if let Some(ctrl) = self.form.controls.iter_mut().find(|c| c.id == *id) {
                    ctrl.properties.set(prop_name, hex.to_string());
                }
            }
        }
    }

    pub fn scroll_properties(&mut self, amount: f32) {
        let items = match self.prop_tab {
            PropTab::Properties => self.collect_props(),
            PropTab::Events => self.collect_events(),
        };
        let total_h = Self::items_height(&items);
        let max_scroll = (total_h - 300.0).max(0.0); // approximate visible height
        self.prop_scroll_y = (self.prop_scroll_y - amount).clamp(0.0, max_scroll);
    }

    fn hit_test_tray(&self, mx: f32, my: f32, rect: Rect) -> Option<Uuid> {
        let form_x = rect.x + FORM_PADDING - self.scroll_x;
        let form_y = rect.y + FORM_PADDING - self.scroll_y;
        let form_h = self.form.height as f32;
        let tray_y = form_y + form_h + 8.0 + 12.0;

        let non_visuals = self.non_visual_controls();
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

    fn is_container(ct: &ControlType) -> bool {
        matches!(ct,
            ControlType::Panel | ControlType::Frame | ControlType::PictureBox |
            ControlType::TabControl | ControlType::SplitContainer |
            ControlType::FlowLayoutPanel | ControlType::TableLayoutPanel
        )
    }

    fn render_controls_recursive(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        parent_id: Option<Uuid>, offset_x: f32, offset_y: f32, s: f32, depth: usize,
    ) {
        if depth > 20 { return; }
        for ctrl in &self.form.controls {
            if ctrl.control_type.is_non_visual() { continue; }
            if ctrl.parent_id != parent_id { continue; }

            self.render_control(pix, fs, sc, ctrl, offset_x, offset_y, s);

            if Self::is_container(&ctrl.control_type) {
                let child_x = offset_x + ctrl.bounds.x as f32;
                let child_y = offset_y + ctrl.bounds.y as f32;
                self.render_controls_recursive(
                    pix, fs, sc, Some(ctrl.id),
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

        let grey = CosmicColor::rgba(150, 150, 150, 255);
        let display_text = if ctrl_text.is_empty() { &ctrl.name } else { &ctrl_text };
        match ctrl.control_type {
            ControlType::Button => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(225, 225, 225, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(173, 173, 173, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                let tw = measure_text_with_font(fs, display_text, font_prop, 12.0, s);
                draw_text_with_font(pix, fs, sc, display_text, cx + (cw - tw) / 2.0, cy + (ch - 14.0) / 2.0, font_prop, 12.0, text_color, s);
            }
            ControlType::Label => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); fill(pix, &paint, cx, cy, cw, ch, s); }
                draw_text_with_font(pix, fs, sc, display_text, cx + 2.0, cy + 2.0, font_prop, 12.0, text_color, s);
            }
            ControlType::LinkLabel => {
                draw_text_with_font(pix, fs, sc, display_text, cx + 2.0, cy + 2.0, font_prop, 12.0, CosmicColor::rgba(0, 102, 204, 255), s);
            }
            ControlType::TextBox | ControlType::MaskedTextBox => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                draw_text_with_font(pix, fs, sc, display_text, cx + 3.0, cy + 3.0, font_prop, 12.0, text_color, s);
            }
            ControlType::RichTextBox => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                // Scrollbar indicator
                paint.set_color_rgba8(230, 230, 230, 255);
                fill(pix, &paint, cx + cw - 14.0, cy, 14.0, ch, s);
                draw_text_with_font(pix, fs, sc, display_text, cx + 3.0, cy + 3.0, font_prop, 12.0, text_color, s);
            }
            ControlType::CheckBox => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); fill(pix, &paint, cx, cy, cw, ch, s); }
                // Checkbox box
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx + 2.0, cy + (ch - 13.0) / 2.0, 13.0, 13.0, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx + 2.0, cy + (ch - 13.0) / 2.0, 13.0, 13.0, s);
                draw_text_with_font(pix, fs, sc, display_text, cx + 20.0, cy + 2.0, font_prop, 12.0, text_color, s);
            }
            ControlType::RadioButton => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); fill(pix, &paint, cx, cy, cw, ch, s); }
                // Radio circle (approximated as small box)
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx + 2.0, cy + (ch - 13.0) / 2.0, 13.0, 13.0, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx + 2.0, cy + (ch - 13.0) / 2.0, 13.0, 13.0, s);
                draw_text_with_font(pix, fs, sc, display_text, cx + 20.0, cy + 2.0, font_prop, 12.0, text_color, s);
            }
            ControlType::ComboBox => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                // Drop arrow
                paint.set_color_rgba8(225, 225, 225, 255);
                fill(pix, &paint, cx + cw - 20.0, cy, 20.0, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx + cw - 20.0, cy, 20.0, ch, s);
                draw_text(pix, fs, sc, "\u{25BC}", cx + cw - 16.0, cy + 4.0, 10.0, CosmicColor::rgba(60, 60, 60, 255), s);
                draw_text_with_font(pix, fs, sc, display_text, cx + 3.0, cy + 3.0, font_prop, 12.0, text_color, s);
            }
            ControlType::ListBox => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                // Scrollbar
                paint.set_color_rgba8(230, 230, 230, 255);
                fill(pix, &paint, cx + cw - 14.0, cy, 14.0, ch, s);
                draw_text_with_font(pix, fs, sc, &ctrl.name, cx + 3.0, cy + 3.0, font_prop, 11.0, grey, s);
            }
            ControlType::PictureBox => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(210, 210, 210, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(160, 160, 160, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                let mut pb = PathBuilder::new();
                pb.move_to(cx * s, cy * s); pb.line_to((cx + cw) * s, (cy + ch) * s);
                pb.move_to((cx + cw) * s, cy * s); pb.line_to(cx * s, (cy + ch) * s);
                if let Some(path) = pb.finish() {
                    paint.set_color_rgba8(180, 180, 180, 255);
                    let mut st = Stroke::default(); st.width = 0.5 * s;
                    pix.stroke_path(&path, &paint, &st, Transform::identity(), None);
                }
            }
            ControlType::ProgressBar => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(230, 230, 230, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(6, 176, 37, 255);
                fill(pix, &paint, cx + 1.0, cy + 1.0, (cw - 2.0) * 0.3, ch - 2.0, s);
                paint.set_color_rgba8(188, 188, 188, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::NumericUpDown => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                // Up/down buttons
                paint.set_color_rgba8(225, 225, 225, 255);
                fill(pix, &paint, cx + cw - 18.0, cy, 18.0, ch, s);
                draw_text(pix, fs, sc, "\u{25B2}", cx + cw - 15.0, cy + 1.0, 8.0, CosmicColor::rgba(60, 60, 60, 255), s);
                draw_text(pix, fs, sc, "\u{25BC}", cx + cw - 15.0, cy + ch / 2.0, 8.0, CosmicColor::rgba(60, 60, 60, 255), s);
                draw_text(pix, fs, sc, "0", cx + 4.0, cy + 3.0, 12.0, text_color, s);
            }
            ControlType::TrackBar => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(240, 240, 240, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                // Track line
                let track_y = cy + ch / 2.0;
                paint.set_color_rgba8(188, 188, 188, 255);
                fill(pix, &paint, cx + 10.0, track_y - 2.0, cw - 20.0, 4.0, s);
                // Thumb
                paint.set_color_rgba8(0, 120, 215, 255);
                fill(pix, &paint, cx + 10.0, track_y - 8.0, 10.0, 16.0, s);
            }
            ControlType::DateTimePicker => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(225, 225, 225, 255);
                fill(pix, &paint, cx + cw - 20.0, cy, 20.0, ch, s);
                draw_text(pix, fs, sc, "\u{1F4C5}", cx + cw - 17.0, cy + 3.0, 11.0, CosmicColor::rgba(60, 60, 60, 255), s);
                draw_text(pix, fs, sc, "1/1/2024", cx + 4.0, cy + 3.0, 11.0, text_color, s);
            }
            ControlType::TreeView | ControlType::ListView => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(230, 230, 230, 255);
                fill(pix, &paint, cx + cw - 14.0, cy, 14.0, ch, s);
                draw_text_with_font(pix, fs, sc, &ctrl.name, cx + 4.0, cy + 4.0, font_prop, 11.0, grey, s);
            }
            ControlType::DataGridView => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                // Header row
                paint.set_color_rgba8(230, 230, 230, 255);
                fill(pix, &paint, cx, cy, cw, 22.0_f32.min(ch), s);
                // Grid lines
                paint.set_color_rgba8(210, 210, 210, 255);
                let col_w = 80.0;
                let mut gx = cx + col_w;
                while gx < cx + cw { fill(pix, &paint, gx, cy, 1.0, ch, s); gx += col_w; }
                let mut gy = cy + 22.0;
                while gy < cy + ch { fill(pix, &paint, cx, gy, cw, 1.0, s); gy += 22.0; }
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::Panel | ControlType::Frame | ControlType::SplitContainer |
            ControlType::FlowLayoutPanel | ControlType::TableLayoutPanel => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(236, 236, 236, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(160, 160, 160, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                if ctrl.control_type == ControlType::Frame {
                    draw_text_with_font(pix, fs, sc, display_text, cx + 8.0, cy + 2.0, font_prop, 11.0, text_color, s);
                } else {
                    draw_text_with_font(pix, fs, sc, &ctrl.name, cx + 4.0, cy + 4.0, font_prop, 11.0, grey, s);
                }
            }
            ControlType::TabControl => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(240, 240, 240, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                // Tab header
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx, cy, 80.0_f32.min(cw), 24.0_f32.min(ch), s);
                paint.set_color_rgba8(160, 160, 160, 255);
                stroke_rect(pix, &paint, cx, cy, 80.0_f32.min(cw), 24.0_f32.min(ch), s);
                draw_text(pix, fs, sc, "TabPage1", cx + 6.0, cy + 4.0, 11.0, text_color, s);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::MenuStrip | ControlType::ToolStrip | ControlType::StatusStrip => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(240, 240, 240, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(204, 204, 204, 255);
                fill(pix, &paint, cx, cy + ch - 1.0, cw, 1.0, s);
                let label = match ctrl.control_type {
                    ControlType::MenuStrip => "File  Edit  View",
                    ControlType::ToolStrip => "[Toolbar]",
                    ControlType::StatusStrip => "Ready",
                    _ => "",
                };
                draw_text(pix, fs, sc, label, cx + 6.0, cy + 3.0, 11.0, text_color, s);
            }
            ControlType::MonthCalendar => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(0, 120, 215, 255);
                fill(pix, &paint, cx, cy, cw, 24.0_f32.min(ch), s);
                draw_text(pix, fs, sc, "January 2024", cx + 6.0, cy + 4.0, 12.0, CosmicColor::rgba(255, 255, 255, 255), s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::WebBrowser => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(255, 255, 255, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                // Address bar
                paint.set_color_rgba8(240, 240, 240, 255);
                fill(pix, &paint, cx, cy, cw, 26.0_f32.min(ch), s);
                paint.set_color_rgba8(255, 255, 255, 255);
                fill(pix, &paint, cx + 4.0, cy + 3.0, cw - 8.0, 20.0, s);
                paint.set_color_rgba8(200, 200, 200, 255);
                stroke_rect(pix, &paint, cx + 4.0, cy + 3.0, cw - 8.0, 20.0, s);
                draw_text(pix, fs, sc, "about:blank", cx + 8.0, cy + 5.0, 11.0, grey, s);
                paint.set_color_rgba8(122, 122, 122, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::HScrollBar => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(230, 230, 230, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(188, 188, 188, 255);
                fill(pix, &paint, cx + 16.0, cy + 2.0, 30.0, ch - 4.0, s);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            ControlType::VScrollBar => {
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(230, 230, 230, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(188, 188, 188, 255);
                fill(pix, &paint, cx + 2.0, cy + 16.0, cw - 4.0, 30.0, s);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
            }
            _ => {
                // Default fallback for any unhandled control type
                if let Some((r, g, b, a)) = back_color { paint.set_color_rgba8(r, g, b, a); } else { paint.set_color_rgba8(230, 230, 230, 255); }
                fill(pix, &paint, cx, cy, cw, ch, s);
                paint.set_color_rgba8(160, 160, 160, 255);
                stroke_rect(pix, &paint, cx, cy, cw, ch, s);
                draw_text_with_font(pix, fs, sc, display_text, cx + 4.0, cy + 4.0, font_prop, 11.0, grey, s);
            }
        }

        if self.selected_controls.contains(&ctrl.id) {
            self.render_selection_handles(pix, cx, cy, cw, ch, s);
        }
    }

    fn render_selection_handles(&self, pix: &mut Pixmap, x: f32, y: f32, w: f32, h: f32, s: f32) {
        let mut paint = Paint::default();
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

    pub fn place_control(&mut self, mx: f32, my: f32, rect: Rect) -> bool {
        let ct = match self.toolbox.selected_tool() {
            ControlTool::Control(ct) => ct,
            ControlTool::Pointer => return false,
        };
        let (cx0, cy0) = self.form_client_origin(rect);
        let parent_id = Self::find_container_at(&self.form, None, cx0, cy0, mx, my, 0);

        let (parent_gx, parent_gy) = if let Some(pid) = parent_id {
            Self::compute_global_pos(&self.form, pid, cx0, cy0)
        } else {
            (cx0, cy0)
        };

        let local_x = snap(((mx - parent_gx).max(0.0)) as i32);
        let local_y = snap(((my - parent_gy).max(0.0)) as i32);

        let name = next_ctrl_name(&ct, &self.form);
        let mut ctrl = Control::new(ct.clone(), name, local_x, local_y);
        ctrl.parent_id = parent_id;
        let id = ctrl.id;
        self.form.controls.push(ctrl);
        self.selected_controls = vec![id];
        true
    }

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
                    if let Some(deeper) = Self::find_container_at(
                        form, Some(ctrl.id), cx, cy, mx, my, depth + 1,
                    ) {
                        return Some(deeper);
                    }
                    return Some(ctrl.id);
                }
            }
        }
        parent_id
    }

    fn hit_test_handle(&self, mx: f32, my: f32, rect: Rect) -> Option<(Uuid, ResizeHandle)> {
        if self.selected_controls.len() != 1 { return None; }
        let id = self.selected_controls[0];
        let ctrl = self.form.controls.iter().find(|c| c.id == id)?;
        if ctrl.control_type.is_non_visual() { return None; }
        let (cx0, cy0) = self.form_client_origin(rect);
        let global = Self::compute_global_pos(&self.form, id, cx0, cy0);
        let x = global.0;
        let y = global.1;
        let w = ctrl.bounds.width as f32;
        let h = ctrl.bounds.height as f32;
        let half = HANDLE_SZ / 2.0 + 2.0;

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

    fn hit_test_controls(
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

    fn compute_global_pos(form: &Form, ctrl_id: Uuid, form_x: f32, form_y: f32) -> (f32, f32) {
        let mut x = form_x;
        let mut y = form_y;
        let mut current_id = Some(ctrl_id);
        let mut offsets = Vec::new();
        while let Some(cid) = current_id {
            if let Some(ctrl) = form.controls.iter().find(|c| c.id == cid) {
                offsets.push((ctrl.bounds.x as f32, ctrl.bounds.y as f32));
                current_id = ctrl.parent_id;
            } else {
                break;
            }
        }
        for (dx, dy) in &offsets {
            x += dx;
            y += dy;
        }
        (x, y)
    }

    pub fn handle_mouse_down(&mut self, mx: f32, my: f32, rect: Rect, ctrl_held: bool) -> bool {
        let lay = self.layout(rect);
        let content_rect = lay.content;
        let toolbox_rect = lay.toolbox;
        let properties_rect = lay.properties;

        // Close menu dropdown when clicking on the form area
        if self.menu_bar.open_menu.is_some() {
            self.menu_bar.open_menu = None;
        }

        if self.toolbox.handle_click(mx, my, toolbox_rect) {
            return true;
        }

        // Properties panel — absorb ALL clicks in properties area
        if properties_rect.contains(mx, my) {
            self.handle_properties_click(mx, my, properties_rect);
            return true;
        }

        // Close color picker when clicking outside the properties panel
        if self.color_picker.open {
            let hex = self.color_picker.color.to_hex();
            if let Some(prop_name) = self.color_picker_prop.take() {
                self.apply_color_prop(&prop_name, &hex);
            }
            self.color_picker.open = false;
            // Don't return — let the click also be processed normally
        }

        if !content_rect.contains(mx, my) { return false; }

        if let Some((id, handle)) = self.hit_test_handle(mx, my, content_rect) {
            if let Some(ctrl) = self.form.controls.iter().find(|c| c.id == id) {
                self.resize_handle = Some(handle);
                self.resize_initial = Some((ctrl.bounds.x, ctrl.bounds.y, ctrl.bounds.width, ctrl.bounds.height));
                self.drag_start = Some((mx, my));
            }
            return true;
        }

        let (cx0, cy0) = self.form_client_origin(content_rect);

        // Try placing control from toolbox
        if self.toolbox.selected_tool() != ControlTool::Pointer {
            let placed = self.place_control(mx, my, content_rect);
            self.toolbox.reset_to_pointer();
            return placed;
        }

        if let Some(hit_id) = Self::hit_test_controls(&self.form, None, cx0, cy0, mx, my, 0) {
            let global = Self::compute_global_pos(&self.form, hit_id, cx0, cy0);
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
            // Store initial bounds of all selected controls for relative dragging
            self.drag_initial_bounds = self.selected_controls.iter().filter_map(|&id| {
                let c = self.form.controls.iter().find(|c| c.id == id)?;
                Some((id, c.bounds.x, c.bounds.y))
            }).collect();
            return true;
        }

        if let Some(id) = self.hit_test_tray(mx, my, content_rect) {
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

        if !ctrl_held {
            self.selected_controls.clear();
        }
        self.lasso_start = Some((mx, my));
        self.lasso_current = Some((mx, my));
        true
    }

    pub fn handle_mouse_move(&mut self, mx: f32, my: f32, rect: Rect) {
        // Route drag to color picker if open
        if self.color_picker.open {
            let lay = self.layout(rect);
            let popup_x = lay.properties.x + 10.0;
            let popup_y = lay.properties.y + PROP_HEADER_H + PROP_TAB_H + 40.0;
            match self.color_picker.handle_drag(mx, my, popup_x, popup_y) {
                ColorPickerEvent::Changed(c) => {
                    let hex = c.to_hex();
                    if let Some(prop_name) = self.color_picker_prop.clone() {
                        self.apply_color_prop(&prop_name, &hex);
                    }
                    return;
                }
                _ => {}
            }
        }

        let _content_rect = self.layout(rect).content;

        if self.lasso_start.is_some() {
            self.lasso_current = Some((mx, my));
            return;
        }

        if let (Some(handle), Some(initial), Some(start)) = (self.resize_handle, self.resize_initial, self.drag_start) {
            let dx = (mx - start.0) as i32;
            let dy = (my - start.1) as i32;
            let (ix, iy, iw, ih) = initial;
            if self.selected_controls.is_empty() { return; }
            let id = self.selected_controls[0];
            if let Some(ctrl) = self.form.controls.iter_mut().find(|c| c.id == id) {
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
                ctrl.bounds.x = snap(nx).max(0);
                ctrl.bounds.y = snap(ny).max(0);
                ctrl.bounds.width = snap(nw).max(10);
                ctrl.bounds.height = snap(nh).max(10);
            }
            return;
        }

        if let Some(start) = self.drag_start {
            let adx = (mx - start.0).abs();
            let ady = (my - start.1).abs();
            if !self.dragging && (adx > 5.0 || ady > 5.0) {
                self.dragging = true;
            }
            if self.dragging {
                let dx = (mx - start.0) as i32;
                let dy = (my - start.1) as i32;
                let initials = self.drag_initial_bounds.clone();
                for (id, ix, iy) in &initials {
                    if let Some(ctrl) = self.form.controls.iter_mut().find(|c| c.id == *id) {
                        ctrl.bounds.x = snap(ix + dx).max(0);
                        ctrl.bounds.y = snap(iy + dy).max(0);
                    }
                }
            }
        }
    }

    pub fn handle_mouse_up(&mut self, rect: Rect) {
        if self.color_picker.open {
            self.color_picker.handle_mouse_up();
        }

        let content_rect = self.layout(rect).content;

        if let (Some(start), Some(end)) = (self.lasso_start, self.lasso_current) {
            let (cx0, cy0) = self.form_client_origin(content_rect);
            let lx = start.0.min(end.0);
            let ly = start.1.min(end.1);
            let lw = (start.0 - end.0).abs();
            let lh = (start.1 - end.1).abs();

            if lw > 3.0 || lh > 3.0 {
                self.selected_controls.clear();
                for ctrl in &self.form.controls {
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
        self.drag_initial_bounds.clear();
        self.resize_handle = None;
        self.resize_initial = None;
    }
}

// Toolbox
struct ToolEntry {
    label: &'static str,
    tool: ControlTool,
}

struct SectionHeader {
    label: &'static str,
}

enum ToolItem {
    Entry(ToolEntry),
    Header(SectionHeader),
}

fn build_items() -> Vec<ToolItem> {
    use ControlType::*;
    vec![
        ToolItem::Entry(ToolEntry { label: "Pointer", tool: ControlTool::Pointer }),
        ToolItem::Entry(ToolEntry { label: "Button", tool: ControlTool::Control(Button) }),
        ToolItem::Entry(ToolEntry { label: "Label", tool: ControlTool::Control(Label) }),
        ToolItem::Entry(ToolEntry { label: "TextBox", tool: ControlTool::Control(TextBox) }),
        ToolItem::Entry(ToolEntry { label: "CheckBox", tool: ControlTool::Control(CheckBox) }),
        ToolItem::Entry(ToolEntry { label: "RadioButton", tool: ControlTool::Control(RadioButton) }),
        ToolItem::Entry(ToolEntry { label: "ComboBox", tool: ControlTool::Control(ComboBox) }),
        ToolItem::Entry(ToolEntry { label: "ListBox", tool: ControlTool::Control(ListBox) }),
        ToolItem::Entry(ToolEntry { label: "GroupBox", tool: ControlTool::Control(Frame) }),
        ToolItem::Entry(ToolEntry { label: "PictureBox", tool: ControlTool::Control(PictureBox) }),
        ToolItem::Entry(ToolEntry { label: "RichTextBox", tool: ControlTool::Control(RichTextBox) }),
        ToolItem::Entry(ToolEntry { label: "WebBrowser", tool: ControlTool::Control(WebBrowser) }),
        ToolItem::Entry(ToolEntry { label: "TreeView", tool: ControlTool::Control(TreeView) }),
        ToolItem::Entry(ToolEntry { label: "DataGridView", tool: ControlTool::Control(DataGridView) }),
        ToolItem::Entry(ToolEntry { label: "Panel", tool: ControlTool::Control(Panel) }),
        ToolItem::Entry(ToolEntry { label: "ListView", tool: ControlTool::Control(ListView) }),
        ToolItem::Entry(ToolEntry { label: "TabControl", tool: ControlTool::Control(TabControl) }),
        ToolItem::Entry(ToolEntry { label: "ProgressBar", tool: ControlTool::Control(ProgressBar) }),
        ToolItem::Entry(ToolEntry { label: "NumericUpDown", tool: ControlTool::Control(NumericUpDown) }),
        ToolItem::Entry(ToolEntry { label: "MenuStrip", tool: ControlTool::Control(MenuStrip) }),
        ToolItem::Entry(ToolEntry { label: "ContextMenuStrip", tool: ControlTool::Control(ContextMenuStrip) }),
        ToolItem::Entry(ToolEntry { label: "StatusStrip", tool: ControlTool::Control(StatusStrip) }),
        ToolItem::Entry(ToolEntry { label: "DateTimePicker", tool: ControlTool::Control(DateTimePicker) }),
        ToolItem::Entry(ToolEntry { label: "LinkLabel", tool: ControlTool::Control(LinkLabel) }),
        ToolItem::Entry(ToolEntry { label: "ToolStrip", tool: ControlTool::Control(ToolStrip) }),
        ToolItem::Entry(ToolEntry { label: "TrackBar", tool: ControlTool::Control(TrackBar) }),
        ToolItem::Entry(ToolEntry { label: "MaskedTextBox", tool: ControlTool::Control(MaskedTextBox) }),
        ToolItem::Entry(ToolEntry { label: "SplitContainer", tool: ControlTool::Control(SplitContainer) }),
        ToolItem::Entry(ToolEntry { label: "FlowLayoutPanel", tool: ControlTool::Control(FlowLayoutPanel) }),
        ToolItem::Entry(ToolEntry { label: "TableLayoutPanel", tool: ControlTool::Control(TableLayoutPanel) }),
        ToolItem::Entry(ToolEntry { label: "MonthCalendar", tool: ControlTool::Control(MonthCalendar) }),
        ToolItem::Entry(ToolEntry { label: "HScrollBar", tool: ControlTool::Control(HScrollBar) }),
        ToolItem::Entry(ToolEntry { label: "VScrollBar", tool: ControlTool::Control(VScrollBar) }),
        ToolItem::Header(SectionHeader { label: "DATA" }),
        ToolItem::Entry(ToolEntry { label: "\u{1F517} BindingSource", tool: ControlTool::Control(BindingSourceComponent) }),
        ToolItem::Entry(ToolEntry { label: "\u{1F9ED} BindingNavigator", tool: ControlTool::Control(BindingNavigator) }),
        ToolItem::Entry(ToolEntry { label: "\u{1F5C4} DataSet", tool: ControlTool::Control(DataSetComponent) }),
        ToolItem::Entry(ToolEntry { label: "\u{1F4CB} DataTable", tool: ControlTool::Control(DataTableComponent) }),
        ToolItem::Entry(ToolEntry { label: "\u{1F50C} DataAdapter", tool: ControlTool::Control(DataAdapterComponent) }),
    ]
}

const ITEM_H: f32 = 22.0;
const HEADER_H: f32 = 26.0;
const SECTION_H: f32 = 20.0;
const SCROLLBAR_W: f32 = 10.0;

pub struct ToolboxState {
    pub selected_idx: Option<usize>,
    pub scroll_y: f32,
}

impl ToolboxState {
    pub fn new() -> Self {
        Self { selected_idx: Some(0), scroll_y: 0.0 }
    }

    pub fn selected_tool(&self) -> ControlTool {
        let items = build_items();
        let mut entry_idx = 0usize;
        for item in &items {
            if let ToolItem::Entry(e) = item {
                if self.selected_idx == Some(entry_idx) {
                    return e.tool.clone();
                }
                entry_idx += 1;
            }
        }
        ControlTool::Pointer
    }

    pub fn reset_to_pointer(&mut self) {
        self.selected_idx = Some(0);
    }

    pub fn scroll(&mut self, amount: f32, rect: Rect) {
        self.scroll_y = (self.scroll_y + amount).clamp(0.0, self.max_scroll(rect));
    }

    fn content_h(&self) -> f32 {
        let items = build_items();
        let mut h = 0.0;
        for item in &items {
            match item {
                ToolItem::Entry(_) => h += ITEM_H,
                ToolItem::Header(_) => h += SECTION_H,
            }
        }
        h
    }

    fn max_scroll(&self, rect: Rect) -> f32 {
        (self.content_h() - (rect.h - HEADER_H)).max(0.0)
    }

    pub fn render(
        &self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache,
        rect: Rect, scale: f32,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        let title_color = CosmicColor::rgba(50, 50, 50, 255);
        draw_text(pix, fs, sc, "Toolbox", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);

        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + HEADER_H - 1.0, rect.w, 1.0, s);
        fill(pix, &paint, rect.x, rect.y, 1.0, rect.h, s); // Left border

        let items = build_items();
        let text_color = CosmicColor::rgba(30, 30, 30, 255);
        let section_color = CosmicColor::rgba(100, 100, 100, 255);
        let list_top = rect.y + HEADER_H;
        let list_h = rect.h - HEADER_H;
        let mut y = list_top - self.scroll_y;
        let mut entry_idx = 0usize;

        for item in &items {
            match item {
                ToolItem::Header(hdr) => {
                    if y + SECTION_H > list_top && y < list_top + list_h {
                        paint.set_color_rgba8(204, 204, 204, 255);
                        fill(pix, &paint, rect.x + 8.0, y + 2.0, rect.w - 16.0, 1.0, s);
                        draw_text(pix, fs, sc, hdr.label, rect.x + 10.0, y + 6.0, 10.0, section_color, s);
                    }
                    y += SECTION_H;
                }
                ToolItem::Entry(e) => {
                    if y + ITEM_H > list_top && y < list_top + list_h {
                        let is_sel = self.selected_idx == Some(entry_idx);
                        if is_sel {
                            paint.set_color_rgba8(0, 120, 212, 255);
                            fill(pix, &paint, rect.x + 4.0, y, rect.w - SCROLLBAR_W - 8.0, ITEM_H, s);
                            let white = CosmicColor::rgba(255, 255, 255, 255);
                            draw_text(pix, fs, sc, e.label, rect.x + 12.0, y + 3.0, 12.0, white, s);
                        } else {
                            draw_text(pix, fs, sc, e.label, rect.x + 12.0, y + 3.0, 12.0, text_color, s);
                        }
                    }
                    entry_idx += 1;
                    y += ITEM_H;
                }
            }
        }

        paint.set_color_rgba8(250, 250, 250, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, HEADER_H, s);
        draw_text(pix, fs, sc, "Toolbox", rect.x + 10.0, rect.y + 6.0, 13.0, title_color, s);
        paint.set_color_rgba8(204, 204, 204, 255);
        fill(pix, &paint, rect.x, rect.y + HEADER_H - 1.0, rect.w, 1.0, s);
    }

    pub fn handle_click(&mut self, mx: f32, my: f32, rect: Rect) -> bool {
        if !rect.contains(mx, my) { return false; }

        let items = build_items();
        let list_top = rect.y + HEADER_H;
        let mut y = list_top - self.scroll_y;
        let mut entry_idx = 0usize;

        for item in &items {
            match item {
                ToolItem::Header(_) => { y += SECTION_H; }
                ToolItem::Entry(_) => {
                    if my >= y && my < y + ITEM_H && my >= list_top {
                        self.selected_idx = Some(entry_idx);
                        return true;
                    }
                    entry_idx += 1;
                    y += ITEM_H;
                }
            }
        }
        false
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
