use std::collections::{HashMap, HashSet};
use std::time::{Duration, Instant};
use std::num::NonZeroU32;
use std::sync::Arc;
use std::fs;
use cosmic_text::{Attrs, Buffer, Color, Family, FontSystem, Metrics, Motion, Shaping, SwashCache, Action, Edit, AttrsList, Cursor, Selection};
use lsp_types::Diagnostic;
use fuzzy_matcher::FuzzyMatcher;
use fuzzy_matcher::skim::SkimMatcherV2;
use tiny_skia::{Color as SkiaColor, Paint, Pixmap, PixmapPaint, Rect, Transform, ColorU8, Stroke, PathBuilder};
use winit::application::ApplicationHandler;
use winit::event::{WindowEvent, ElementState, MouseButton, MouseScrollDelta};
use winit::event_loop::{ControlFlow, EventLoop};
use winit::window::{Window, WindowAttributes};
use winit::keyboard::{Key, NamedKey};
#[cfg(target_os = "macos")]
use winit::platform::modifier_supplement::KeyEventExtModifierSupplement;
use softbuffer::{Context, Surface};
use arboard::Clipboard;

use std::path::Path;
use serde::{Deserialize, Serialize};

use crate::editor::{Editor as MyEditor, TokenKind};
use crate::language::{load_language, LanguageDef};
use crate::lsp_client::{LspClient, LspRequest, LspEvent};
use vybe_widgets::{TreeView, TreeEvent, Dropdown, DropdownEvent};
use vybe_widgets::code_editor_widget::{Theme, CodeEditorWidget, apply_highlighting};

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct Keybinding {
    pub key: String,
    #[serde(default)] pub cmd: bool,
    #[serde(default)] pub shift: bool,
    #[serde(default)] pub alt: bool,
    pub action: String,
}

const SCALE: f32 = 2.0;
const EXPLORER_WIDTH: f32 = 250.0; 
const TAB_BAR_HEIGHT: f32 = 36.0;
const MINIMAP_WIDTH: f32 = 80.0;
const UI_BAR_HEIGHT: f32 = 0.0;
const FOOTER_HEIGHT: f32 = 24.0;
const GUTTER_WIDTH: f32 = 64.0;
const SPLITTER_WIDTH: f32 = 4.0;
const SIDEBAR_TAB_H: f32 = 28.0;

#[derive(Clone, Copy, PartialEq)]
enum SidebarTab { Files, Project }

struct ProjectExplorerState {
    scroll_y: f32,
    forms_collapsed: bool,
    code_collapsed: bool,
    refs_collapsed: bool,
    resources_collapsed: bool,
}

impl ProjectExplorerState {
    fn new() -> Self { Self { scroll_y: 0.0, forms_collapsed: false, code_collapsed: false, refs_collapsed: false, resources_collapsed: false } }
}

struct ProjectPropsDialog {
    visible: bool,
    selected_startup: usize,
}

impl ProjectPropsDialog {
    fn new() -> Self { Self { visible: false, selected_startup: 0 } }

    fn open(&mut self, project: &vybe_project::project::Project) {
        self.visible = true;
        self.selected_startup = match &project.startup_object {
            vybe_project::project::StartupObject::SubMain => 0,
            vybe_project::project::StartupObject::None => 1,
            vybe_project::project::StartupObject::Form(name) => {
                project.forms.iter().position(|f| &f.form.name == name).map(|i| i + 2).unwrap_or(0)
            }
        };
    }

    fn close(&mut self) { self.visible = false; }

    fn apply(&self, project: &mut vybe_project::project::Project) {
        match self.selected_startup {
            0 => { project.startup_object = vybe_project::project::StartupObject::SubMain; project.startup_form = None; }
            1 => { project.startup_object = vybe_project::project::StartupObject::None; project.startup_form = None; }
            n => {
                let idx = n - 2;
                if let Some(fm) = project.forms.get(idx) {
                    let name = fm.form.name.clone();
                    project.startup_object = vybe_project::project::StartupObject::Form(name.clone());
                    project.startup_form = Some(name);
                }
            }
        }
    }

    fn startup_options(&self, project: &vybe_project::project::Project) -> Vec<String> {
        let mut opts = vec!["Sub Main".into(), "(None)".into()];
        for fm in &project.forms { opts.push(fm.form.name.clone()); }
        opts
    }

    fn render(&self, pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, win_w: f32, win_h: f32, scale: f32, project: &vybe_project::project::Project) {
        if !self.visible { return; }
        let s = scale;
        let mut paint = Paint::default();
        let dw = 400.0f32; let dh = 280.0f32;
        let dx = (win_w - dw) / 2.0; let dy = (win_h - dh) / 2.0;

        // Overlay
        paint.set_color_rgba8(0, 0, 0, 120);
        if let Some(r) = tiny_skia::Rect::from_xywh(0.0, 0.0, win_w * s, win_h * s) { pix.fill_rect(r, &paint, Transform::identity(), None); }

        // Shadow + bg
        paint.set_color_rgba8(0, 0, 0, 40);
        if let Some(r) = tiny_skia::Rect::from_xywh((dx+4.0)*s, (dy+4.0)*s, dw*s, dh*s) { pix.fill_rect(r, &paint, Transform::identity(), None); }
        paint.set_color_rgba8(255, 255, 255, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(dx*s, dy*s, dw*s, dh*s) { pix.fill_rect(r, &paint, Transform::identity(), None); }

        // Title bar
        paint.set_color_rgba8(0, 120, 212, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(dx*s, dy*s, dw*s, 30.0*s) { pix.fill_rect(r, &paint, Transform::identity(), None); }
        let white = Color::rgba(255, 255, 255, 255);
        let text_col = Color::rgba(30, 30, 30, 255);
        crate::ide_text::draw_text(pix, fs, sc, &format!("{} - Project Properties", project.name), dx + 10.0, dy + 7.0, 13.0, white, s);
        crate::ide_text::draw_text(pix, fs, sc, "X", dx + dw - 22.0, dy + 7.0, 13.0, white, s);

        let label_col = Color::rgba(60, 60, 60, 255);
        let mut y = dy + 30.0 + 16.0;

        // Project Name
        crate::ide_text::draw_text(pix, fs, sc, "Project Name:", dx + 16.0, y, 12.0, label_col, s);
        y += 20.0;
        paint.set_color_rgba8(245, 245, 245, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh((dx+16.0)*s, y*s, (dw-32.0)*s, 28.0*s) { pix.fill_rect(r, &paint, Transform::identity(), None); }
        crate::ide_text::draw_text(pix, fs, sc, &project.name, dx + 22.0, y + 6.0, 12.0, text_col, s);
        y += 28.0 + 16.0;

        // Startup Object
        crate::ide_text::draw_text(pix, fs, sc, "Startup Object:", dx + 16.0, y, 12.0, label_col, s);
        y += 20.0;

        let options = self.startup_options(project);
        let opt_h = 22.0f32;
        for (i, label) in options.iter().enumerate() {
            let oy = y + i as f32 * opt_h;
            if i == self.selected_startup {
                paint.set_color_rgba8(0, 120, 212, 255);
                if let Some(r) = tiny_skia::Rect::from_xywh((dx+17.0)*s, (oy+1.0)*s, (dw-34.0)*s, (opt_h-2.0)*s) { pix.fill_rect(r, &paint, Transform::identity(), None); }
                crate::ide_text::draw_text(pix, fs, sc, label, dx + 24.0, oy + 4.0, 11.0, white, s);
            } else {
                crate::ide_text::draw_text(pix, fs, sc, label, dx + 24.0, oy + 4.0, 11.0, text_col, s);
            }
        }

        // Footer
        let footer_y = dy + dh - 44.0;
        paint.set_color_rgba8(240, 240, 240, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(dx*s, footer_y*s, dw*s, 44.0*s) { pix.fill_rect(r, &paint, Transform::identity(), None); }

        // OK button
        let btn_w = 80.0; let btn_h = 28.0;
        let ok_x = dx + dw - btn_w * 2.0 - 24.0;
        let btn_y = footer_y + 8.0;
        paint.set_color_rgba8(0, 120, 212, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(ok_x*s, btn_y*s, btn_w*s, btn_h*s) { pix.fill_rect(r, &paint, Transform::identity(), None); }
        crate::ide_text::draw_text(pix, fs, sc, "OK", ok_x + 30.0, btn_y + 6.0, 12.0, white, s);

        // Cancel button
        let cancel_x = dx + dw - btn_w - 16.0;
        paint.set_color_rgba8(240, 240, 240, 255);
        if let Some(r) = tiny_skia::Rect::from_xywh(cancel_x*s, btn_y*s, btn_w*s, btn_h*s) { pix.fill_rect(r, &paint, Transform::identity(), None); }
        crate::ide_text::draw_text(pix, fs, sc, "Cancel", cancel_x + 20.0, btn_y + 6.0, 12.0, text_col, s);
    }

    fn handle_click(&mut self, mx: f32, my: f32, win_w: f32, win_h: f32, project: &vybe_project::project::Project) -> bool {
        if !self.visible { return false; }
        let dw = 400.0f32; let dh = 280.0f32;
        let dx = (win_w - dw) / 2.0; let dy = (win_h - dh) / 2.0;

        // Outside dialog
        if mx < dx || mx > dx + dw || my < dy || my > dy + dh { self.close(); return true; }

        // Close button
        if mx >= dx + dw - 28.0 && my < dy + 30.0 { self.close(); return true; }

        // Startup option list
        let list_y = dy + 30.0 + 16.0 + 20.0 + 28.0 + 16.0 + 20.0;
        let options = self.startup_options(project);
        let opt_h = 22.0;
        if mx >= dx + 16.0 && mx < dx + dw - 16.0 {
            let rel_y = my - list_y;
            if rel_y >= 0.0 {
                let idx = (rel_y / opt_h) as usize;
                if idx < options.len() { self.selected_startup = idx; return true; }
            }
        }

        // Cancel button
        let footer_y = dy + dh - 44.0;
        let btn_w = 80.0; let btn_h = 28.0;
        let cancel_x = dx + dw - btn_w - 16.0;
        let btn_y = footer_y + 8.0;
        if mx >= cancel_x && mx < cancel_x + btn_w && my >= btn_y && my < btn_y + btn_h {
            self.close(); return true;
        }

        true // consume click inside dialog
    }

    fn is_ok_clicked(&self, mx: f32, my: f32, win_w: f32, win_h: f32) -> bool {
        if !self.visible { return false; }
        let dw = 400.0f32; let dh = 280.0f32;
        let dx = (win_w - dw) / 2.0; let dy = (win_h - dh) / 2.0;
        let footer_y = dy + dh - 44.0;
        let btn_w = 80.0; let btn_h = 28.0;
        let ok_x = dx + dw - btn_w * 2.0 - 24.0;
        let btn_y = footer_y + 8.0;
        mx >= ok_x && mx < ok_x + btn_w && my >= btn_y && my < btn_y + btn_h
    }
}

pub enum TabContent {
    Code(CodeEditorWidget),
    Form(crate::form_designer_tab::FormDesignerState),
    Resources(vybe_widgets::ResourceEditor),
}

pub struct Tab {
    pub name: String,
    pub path: Option<String>,
    pub content: TabContent,
    pub is_sticky: bool,
    pub buffer: Option<Buffer>,
    pub is_modified: bool,
}

// Theme, CachedGlyph, apply_highlighting, and CodeEditorWidget are now in vybe_widgets::code_editor_widget.

// ── Below was Theme + CodeEditorWidget (now in vybe_widgets) ──
// Removed ~1100 lines. See vybe_widgets::code_editor_widget.

struct App {
    window: Option<Arc<Window>>,
    context: Option<Context<Arc<Window>>>,
    surface: Option<Surface<Arc<Window>, Arc<Window>>>,
    font_system: FontSystem,
    swash_cache: SwashCache,
    pixmap: Option<Pixmap>,
    tabs: Vec<Tab>,
    active_tab: usize,
    tree_view: TreeView,
    all_languages: Vec<String>,
    current_lang: String,
    lang_dropdown: Option<Dropdown>,
    theme_dropdown: Option<Dropdown>,
    clipboard: Option<Clipboard>,
    modifiers: winit::event::Modifiers,
    last_click_time: Instant,
    click_count: u32,
    mouse_pos: (f32, f32),
    explorer_width: f32,
    is_dragging_splitter: bool,
    hovering_splitter: bool,
    hovering_tab_close: Option<usize>,
    is_dragging: bool,
    needs_redraw: bool,
    last_lsp_update: Instant,
    pending_lsp_update: bool,
    lsp: Arc<LspClient>,
    is_quick_open: bool,
    quick_open_query: String,
    tab_scroll_x: f32,
    current_theme_idx: usize,
    breadcrumb_rects: Vec<(Rect, String)>,
    keybindings: Vec<Keybinding>,
    open_form: bool,
    sidebar_tab: SidebarTab,
    project: vybe_project::project::Project,
    project_explorer: ProjectExplorerState,
    project_props_dialog: ProjectPropsDialog,
    control_clipboard: Vec<vybe_forms::Control>,
    project_path: Option<String>,
    run_child: Option<std::process::Child>,
    /// Context menu for project explorer: (x, y, item_name)
    pe_context_menu: Option<(f32, f32, String)>,
}

#[derive(Clone, Copy)]
enum EditAction { Undo, Redo, Cut, Copy, Paste, Delete }

impl App {
    fn new(_my_editor: MyEditor, open_form: bool) -> Self {
        let mut langs = Vec::new();
        // Try multiple paths to find the basic-languages folder
        let mut candidates: Vec<std::path::PathBuf> = vec![
            std::path::PathBuf::from("crates/code_editor/basic-languages"),
            std::path::PathBuf::from("basic-languages"),
            std::path::PathBuf::from("../code_editor/basic-languages"),
        ];
        // Walk up from exe to find workspace
        if let Ok(exe) = std::env::current_exe() {
            let mut dir_opt = exe.parent();
            while let Some(dir) = dir_opt {
                let cand = dir.join("crates").join("code_editor").join("basic-languages");
                if cand.exists() { candidates.insert(0, cand); break; }
                dir_opt = dir.parent();
            }
        }
        for path in &candidates {
            if let Ok(es) = std::fs::read_dir(path) {
                for e in es.flatten() {
                    if e.file_type().map(|t| t.is_dir()).unwrap_or(false) {
                        if let Some(n) = e.file_name().to_str() { langs.push(n.to_string()); }
                    }
                }
                if !langs.is_empty() { break; }
            }
        }
        if langs.is_empty() {
            langs = vec!["rust".into(), "javascript".into(), "typescript".into(), "python".into(), "vb".into(), "csharp".into(), "text".into()];
        }
        langs.sort();
        let root_dir = std::env::current_dir().unwrap_or_else(|_| std::path::PathBuf::from("/"));
        let _root_uri = format!("file://{}", root_dir.to_string_lossy());
        Self { 
            window: None, context: None, surface: None, font_system: FontSystem::new(), swash_cache: SwashCache::new(), pixmap: None, tabs: Vec::new(), active_tab: 0, 
            tree_view: TreeView::new(".", 2.0), all_languages: langs, current_lang: "rust".to_string(), lang_dropdown: None, theme_dropdown: None, clipboard: Clipboard::new().ok(), 
            modifiers: winit::event::Modifiers::default(), last_click_time: Instant::now(), click_count: 0, mouse_pos: (0.0, 0.0), 
            explorer_width: EXPLORER_WIDTH, is_dragging_splitter: false, hovering_splitter: false, hovering_tab_close: None,
            is_dragging: false, needs_redraw: true,
            last_lsp_update: Instant::now(), pending_lsp_update: false,
            lsp: Arc::new(LspClient::new()),
            is_quick_open: false,
            quick_open_query: String::new(),
            tab_scroll_x: 0.0,
            current_theme_idx: 0,
            breadcrumb_rects: Vec::new(),
            open_form,
            keybindings: {
                let paths = ["keybindings.json", "crates/code_editor/keybindings.json"];
                let mut kb = Vec::new();
                for p in paths {
                    if let Ok(s) = fs::read_to_string(p) {
                        if let Ok(parsed) = serde_json::from_str::<Vec<Keybinding>>(&s) {
                            kb = parsed;
                            break;
                        }
                    }
                }
                if kb.is_empty() {
                    kb = vec![
                        Keybinding { key: "z".into(), cmd: true, shift: false, alt: false, action: "Undo".into() },
                        Keybinding { key: "z".into(), cmd: true, shift: true, alt: false, action: "Redo".into() },
                        Keybinding { key: "a".into(), cmd: true, shift: false, alt: false, action: "SelectAll".into() },
                        Keybinding { key: "s".into(), cmd: true, shift: false, alt: false, action: "Save".into() },
                        Keybinding { key: "f".into(), cmd: true, shift: false, alt: false, action: "Find".into() },
                        Keybinding { key: "h".into(), cmd: true, shift: false, alt: false, action: "Replace".into() },
                        Keybinding { key: "/".into(), cmd: true, shift: false, alt: false, action: "ToggleComment".into() },
                        Keybinding { key: "Tab".into(), cmd: false, shift: false, alt: false, action: "Indent".into() },
                        Keybinding { key: "Tab".into(), cmd: false, shift: true, alt: false, action: "Unindent".into() },
                        Keybinding { key: "ArrowUp".into(), cmd: true, shift: false, alt: false, action: "MoveBufferStart".into() },
                        Keybinding { key: "ArrowDown".into(), cmd: true, shift: false, alt: false, action: "MoveBufferEnd".into() },
                        Keybinding { key: "ArrowLeft".into(), cmd: true, shift: false, alt: false, action: "MoveLineStart".into() },
                        Keybinding { key: "ArrowRight".into(), cmd: true, shift: false, alt: false, action: "MoveLineEnd".into() },
                        Keybinding { key: "ArrowLeft".into(), cmd: false, shift: false, alt: true, action: "MoveWordLeft".into() },
                        Keybinding { key: "ArrowRight".into(), cmd: false, shift: false, alt: true, action: "MoveWordRight".into() },
                    ];
                }
                kb
            },
            sidebar_tab: SidebarTab::Project,
            project: {
                let mut p = vybe_project::project::Project::new("Project1".to_string());
                let mut form = vybe_forms::Form::new("Form1".to_string());
                form.width = 640; form.height = 480;
                p.forms.push(vybe_project::project::FormModule::new_classic(form));
                p.startup_object = vybe_project::project::StartupObject::Form("Form1".to_string());
                p
            },
            project_explorer: ProjectExplorerState::new(),
            project_props_dialog: ProjectPropsDialog::new(),
            control_clipboard: Vec::new(),
            project_path: None,
            run_child: None,
            pe_context_menu: None,
        }
    }
    /// Height in logical pixels of menu bar + toolbar (always present if any Form tab exists).
    fn top_chrome_h(&self) -> f32 {
        if self.tabs.iter().any(|t| matches!(&t.content, TabContent::Form(_))) {
            28.0 + 36.0
        } else {
            0.0
        }
    }
    pub fn active_theme(&self) -> Theme {
        match self.current_theme_idx {
            0 => Theme::silicon_green(),
            1 => Theme::cloud_blue(),
            2 => Theme::coffee_cream(),
            3 => Theme::sakura_pink(),
            4 => Theme::one_dark(),
            5 => Theme::monokai(),
            6 => Theme::frost_light(),
            7 => Theme::solarized_light(),
            8 => Theme::midnight(),
            9 => Theme::aura(),
            10 => Theme::veridian(),
            11 => Theme::rose(),
            12 => Theme::cyber(),
            13 => Theme::titanium(),
            14 => Theme::indigo_night(),
            _ => Theme::one_dark(),
        }
    }

    pub fn get_theme_name(&self) -> &str {
        match self.current_theme_idx { 
            0 => "Silicon Green", 1 => "Cloud Blue", 2 => "Coffee Cream", 3 => "Sakura Pink", 
            4 => "One Dark", 5 => "Monokai", 6 => "Frost Light", 7 => "Solarized Light", 
            8 => "Midnight", 9 => "Aura", 10 => "Veridian", 11 => "Rose",
            12 => "Cyber", 13 => "Titanium", 14 => "Indigo Night", _ => "One Dark" 
        }
    }

    fn render(&mut self) {
        // Debounce LSP Update
        if self.pending_lsp_update && self.last_lsp_update.elapsed().as_millis() > 300 {
            let mut lsp_text = None;
            if let Some(tab) = self.tabs.get(self.active_tab) {
                if let TabContent::Code(cw) = &tab.content {
                    let text = cw.my_editor.rope.to_string();
                    let uri = tab.path.clone().unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", tab.name));
                    lsp_text = Some((text, uri));
                }
            }
            if let Some((text, uri)) = lsp_text {
                self.lsp.send(LspRequest::Change(text, uri));
            }
            self.pending_lsp_update = false;
        }

        let theme = self.active_theme();
        let theme_name = self.get_theme_name().to_string();
        let (surf, pix) = match (&mut self.surface, &mut self.pixmap) { (Some(s), Some(p)) => (s, p), _ => return };
        
        let to_skia = |c: Color| SkiaColor::from_rgba8(c.r(), c.g(), c.b(), c.a());
        pix.fill(to_skia(theme.bg));

        // Compute top chrome height (menu + toolbar when any Form tab exists)
        let has_form_tab = self.tabs.iter().any(|t| matches!(&t.content, TabContent::Form(_)));
        let top_chrome_h: f32 = if has_form_tab { 28.0 + 36.0 } else { 0.0 };
        let top_chrome_px = top_chrome_h * SCALE;

        // 0. Menu bar + Toolbar (always present when a Form tab exists)
        if has_form_tab {
            // Find the Form tab to get menu_bar state
            if let Some(form_tab) = self.tabs.iter().find(|t| matches!(&t.content, TabContent::Form(_))) {
                if let TabContent::Form(f) = &form_tab.content {
                    let menu_rect = crate::form_designer_tab::Rect { x: 0.0, y: 0.0, w: pix.width() as f32 / SCALE, h: 28.0 };
                    let tb_rect = crate::form_designer_tab::Rect { x: 0.0, y: 28.0, w: pix.width() as f32 / SCALE, h: 36.0 };
                    f.menu_bar.render(pix, &mut self.font_system, &mut self.swash_cache, menu_rect, SCALE);
                    crate::form_designer_tab::render_toolbar_pub(pix, &mut self.font_system, &mut self.swash_cache, tb_rect, SCALE);
                }
            }
        }

        // 1. Sidebar
        let sidebar_x = 0.0;
        let sidebar_top = top_chrome_px;
        let sidebar_w = self.explorer_width * SCALE;
        let sidebar_h = pix.height() as f32 - top_chrome_px;

        // Sidebar background
        let mut sp = Paint::default(); sp.set_color_rgba8(theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), theme.sidebar_bg.a());
        pix.fill_rect(Rect::from_xywh(sidebar_x, sidebar_top, sidebar_w, sidebar_h).unwrap(), &sp, Transform::identity(), None);

        // Sidebar tabs (Files | Project)
        let stab_h = SIDEBAR_TAB_H * SCALE;
        let stab_w = sidebar_w / 2.0;
        let stab_y = sidebar_top;
        // Files tab
        {
            let active = self.sidebar_tab == SidebarTab::Files;
            let mut tp = Paint::default();
            if active { tp.set_color_rgba8(theme.active_tab_bg.r(), theme.active_tab_bg.g(), theme.active_tab_bg.b(), theme.active_tab_bg.a()); }
            else { tp.set_color_rgba8(theme.inactive_tab_bg.r(), theme.inactive_tab_bg.g(), theme.inactive_tab_bg.b(), theme.inactive_tab_bg.a()); }
            pix.fill_rect(Rect::from_xywh(sidebar_x, stab_y, stab_w, stab_h).unwrap(), &tp, Transform::identity(), None);
            if active {
                let mut up = Paint::default(); up.set_color_rgba8(theme.kw.r(), theme.kw.g(), theme.kw.b(), 255);
                pix.fill_rect(Rect::from_xywh(sidebar_x, stab_y + stab_h - 2.0 * SCALE, stab_w, 2.0 * SCALE).unwrap(), &up, Transform::identity(), None);
            }
            let col = if active { theme.active_tab_text } else { theme.inactive_tab_text };
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "Files", sidebar_x + 10.0 * SCALE, stab_y + 6.0 * SCALE, col);
        }
        // Project tab
        {
            let active = self.sidebar_tab == SidebarTab::Project;
            let mut tp = Paint::default();
            if active { tp.set_color_rgba8(theme.active_tab_bg.r(), theme.active_tab_bg.g(), theme.active_tab_bg.b(), theme.active_tab_bg.a()); }
            else { tp.set_color_rgba8(theme.inactive_tab_bg.r(), theme.inactive_tab_bg.g(), theme.inactive_tab_bg.b(), theme.inactive_tab_bg.a()); }
            pix.fill_rect(Rect::from_xywh(sidebar_x + stab_w, stab_y, stab_w, stab_h).unwrap(), &tp, Transform::identity(), None);
            if active {
                let mut up = Paint::default(); up.set_color_rgba8(theme.kw.r(), theme.kw.g(), theme.kw.b(), 255);
                pix.fill_rect(Rect::from_xywh(sidebar_x + stab_w, stab_y + stab_h - 2.0 * SCALE, stab_w, 2.0 * SCALE).unwrap(), &up, Transform::identity(), None);
            }
            let col = if active { theme.active_tab_text } else { theme.inactive_tab_text };
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "Project", sidebar_x + stab_w + 10.0 * SCALE, stab_y + 6.0 * SCALE, col);
        }
        // Tab separator
        {
            let mut lp = Paint::default(); lp.set_color_rgba8(theme.splitter_bg.r(), theme.splitter_bg.g(), theme.splitter_bg.b(), 255);
            pix.fill_rect(Rect::from_xywh(sidebar_x, stab_y + stab_h - 1.0, sidebar_w, 1.0).unwrap(), &lp, Transform::identity(), None);
        }

        // Sidebar content below tabs
        let sidebar_content_y = stab_y + stab_h;
        match self.sidebar_tab {
            SidebarTab::Files => {
                // Sync Sidebar Selection
                if let Some(tab) = self.tabs.get(self.active_tab) {
                    if let Some(path) = &tab.path {
                        self.tree_view.reveal_path(path);
                    }
                }
                // Tree view starts below sidebar tabs (not below editor tab bar)
                self.tree_view.render(pix, &mut self.font_system, &mut self.swash_cache, sidebar_x, sidebar_content_y, sidebar_w, theme.sidebar_text, (theme.selection.r(), theme.selection.g(), theme.selection.b(), theme.selection.a()));
            }
            SidebarTab::Project => {
                let pe = &self.project_explorer;
                let project = &self.project;
                let current_form: Option<&str> = if self.active_tab < self.tabs.len() {
                    if let TabContent::Form(f) = &self.tabs[self.active_tab].content { Some(&f.form.name) } else { None }
                } else { None };
                let pe_x = sidebar_x / SCALE;
                let pe_y = sidebar_content_y / SCALE;
                let pe_w = self.explorer_width;
                let pe_h = (sidebar_h - stab_h) / SCALE;
                let item_h = 24.0f32;
                let indent = 16.0f32;
                let sel_bg = (theme.selection.r(), theme.selection.g(), theme.selection.b(), 80u8);
                let text_col = Color::rgba(theme.sidebar_text.r(), theme.sidebar_text.g(), theme.sidebar_text.b(), theme.sidebar_text.a());
                let dim_col = Color::rgba(theme.sidebar_text.r().saturating_sub(60), theme.sidebar_text.g().saturating_sub(60), theme.sidebar_text.b().saturating_sub(60), 255);
                let mut iy = pe_y - pe.scroll_y;
                let mut pp = Paint::default();

                // Project name
                if iy + item_h > pe_y && iy < pe_y + pe_h {
                    crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("\u{1F4C1} {}", project.name), pe_x + 8.0, iy + 4.0, 12.0, text_col, SCALE);
                }
                iy += item_h;

                // Forms section
                let forms_arrow = if pe.forms_collapsed { "\u{25B6}" } else { "\u{25BC}" };
                if iy + item_h > pe_y && iy < pe_y + pe_h {
                    crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("{} Forms", forms_arrow), pe_x + 8.0 + indent, iy + 4.0, 12.0, text_col, SCALE);
                }
                iy += item_h;

                if !pe.forms_collapsed {
                    for fm in &project.forms {
                        if iy + item_h > pe_y && iy < pe_y + pe_h {
                            let is_sel = current_form == Some(fm.form.name.as_str());
                            if is_sel {
                                pp.set_color_rgba8(sel_bg.0, sel_bg.1, sel_bg.2, sel_bg.3);
                                if let Some(r) = tiny_skia::Rect::from_xywh(pe_x * SCALE, iy * SCALE, pe_w * SCALE, item_h * SCALE) {
                                    pix.fill_rect(r, &pp, Transform::identity(), None);
                                }
                            }
                            crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("  {}", fm.form.name), pe_x + 8.0 + indent * 2.0, iy + 4.0, 12.0, text_col, SCALE);
                        }
                        iy += item_h;
                    }
                }

                // Code section
                if !project.code_files.is_empty() {
                    let code_arrow = if pe.code_collapsed { "\u{25B6}" } else { "\u{25BC}" };
                    if iy + item_h > pe_y && iy < pe_y + pe_h {
                        crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("{} Code", code_arrow), pe_x + 8.0 + indent, iy + 4.0, 12.0, text_col, SCALE);
                    }
                    iy += item_h;
                    if !pe.code_collapsed {
                        for cf in &project.code_files {
                            if iy + item_h > pe_y && iy < pe_y + pe_h {
                                crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("  {}", cf.name), pe_x + 8.0 + indent * 2.0, iy + 4.0, 12.0, text_col, SCALE);
                            }
                            iy += item_h;
                        }
                    }
                }

                // References section
                if !project.project_references.is_empty() {
                    let refs_arrow = if pe.refs_collapsed { "\u{25B6}" } else { "\u{25BC}" };
                    if iy + item_h > pe_y && iy < pe_y + pe_h {
                        crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("{} References", refs_arrow), pe_x + 8.0 + indent, iy + 4.0, 12.0, text_col, SCALE);
                    }
                    iy += item_h;
                    if !pe.refs_collapsed {
                        for rn in &project.project_references {
                            if iy + item_h > pe_y && iy < pe_y + pe_h {
                                crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("  {}", rn), pe_x + 8.0 + indent * 2.0, iy + 4.0, 12.0, dim_col, SCALE);
                            }
                            iy += item_h;
                        }
                    }
                }

                // Resources section — only show if project has resource files with content
                {
                    let has_any_resources = !project.resource_files.is_empty() &&
                        project.resource_files.iter().any(|rm| !rm.resources.is_empty() || rm.file_path.is_some());
                    // Also show if a Resources tab is open
                    let has_res_tab = self.tabs.iter().any(|t| matches!(&t.content, TabContent::Resources(_)));
                    if has_any_resources || has_res_tab {
                        let res_count: usize = project.resource_files.iter().map(|rm| rm.resources.len()).sum();
                        let res_arrow = if pe.resources_collapsed { "\u{25B6}" } else { "\u{25BC}" };
                        let res_label = if res_count > 0 {
                            format!("{} Resources ({})", res_arrow, res_count)
                        } else {
                            format!("{} Resources", res_arrow)
                        };
                        if iy + item_h > pe_y && iy < pe_y + pe_h {
                            crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &res_label, pe_x + 8.0 + indent, iy + 4.0, 12.0, text_col, SCALE);
                        }
                        iy += item_h;
                        if !pe.resources_collapsed {
                            for (_ri, rm) in project.resource_files.iter().enumerate() {
                                let rm_label = format!("  {}.resx ({})", rm.name, rm.resources.len());
                                let is_res_tab = self.tabs.get(self.active_tab).map(|t| matches!(&t.content, TabContent::Resources(_))).unwrap_or(false);
                                if iy + item_h > pe_y && iy < pe_y + pe_h {
                                    if is_res_tab {
                                        pp.set_color_rgba8(sel_bg.0, sel_bg.1, sel_bg.2, sel_bg.3);
                                        if let Some(r) = tiny_skia::Rect::from_xywh(pe_x * SCALE, iy * SCALE, pe_w * SCALE, item_h * SCALE) {
                                            pix.fill_rect(r, &pp, Transform::identity(), None);
                                        }
                                    }
                                    crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &rm_label, pe_x + 8.0 + indent * 2.0, iy + 4.0, 12.0, text_col, SCALE);
                                }
                                iy += item_h;
                            }
                        }
                    }
                }
            }
        }

        // 1b. Splitter
        let mut slp = Paint::default();
        if self.is_dragging_splitter { slp.set_color_rgba8(0,122,204,255); }
        else if self.hovering_splitter { slp.set_color_rgba8(theme.splitter_bg.r(), theme.splitter_bg.g(), theme.splitter_bg.b(), 255); }
        else { slp.set_color_rgba8(theme.splitter_bg.r(), theme.splitter_bg.g(), theme.splitter_bg.b(), 255); }
        pix.fill_rect(Rect::from_xywh(self.explorer_width * SCALE, top_chrome_px, SPLITTER_WIDTH * SCALE, pix.height() as f32 - top_chrome_px).unwrap(), &slp, Transform::identity(), None);

        // Splitter separation line
        let mut lp = Paint::default(); lp.set_color_rgba8(theme.splitter_bg.r().saturating_add(20), theme.splitter_bg.g().saturating_add(20), theme.splitter_bg.b().saturating_add(20), 255);
        pix.fill_rect(Rect::from_xywh((self.explorer_width + SPLITTER_WIDTH) * SCALE, top_chrome_px, 1.0 * SCALE, pix.height() as f32 - top_chrome_px).unwrap(), &lp, Transform::identity(), None);

        // 2. Tab Bar
        let ed_start_x = (self.explorer_width + SPLITTER_WIDTH + 1.0) * SCALE;
        let mut tp = Paint::default(); tp.set_color_rgba8(theme.tab_bar_bg.r(), theme.tab_bar_bg.g(), theme.tab_bar_bg.b(), theme.tab_bar_bg.a());
        pix.fill_rect(Rect::from_xywh(ed_start_x, top_chrome_px, pix.width() as f32 - ed_start_x, TAB_BAR_HEIGHT * SCALE).unwrap(), &tp, Transform::identity(), None);

            let mut tx_off = ed_start_x + self.tab_scroll_x;
            for i in 0..self.tabs.len() {
                if tx_off + 160.0 * SCALE < ed_start_x { tx_off += 160.0 * SCALE; continue; }
                if tx_off > pix.width() as f32 { break; }
                
                let active = i == self.active_tab;
                let tw = 160.0 * SCALE;
                
                // Render Background & Underline
                if active {
                    let mut ap = Paint::default(); ap.set_color_rgba8(theme.active_tab_bg.r(), theme.active_tab_bg.g(), theme.active_tab_bg.b(), theme.active_tab_bg.a());
                    pix.fill_rect(Rect::from_xywh(tx_off, top_chrome_px, tw, TAB_BAR_HEIGHT * SCALE).unwrap(), &ap, Transform::identity(), None);

                    let mut up = Paint::default(); up.set_color_rgba8(theme.kw.r(), theme.kw.g(), theme.kw.b(), 255);
                    pix.fill_rect(Rect::from_xywh(tx_off, top_chrome_px + (TAB_BAR_HEIGHT - 2.0) * SCALE, tw, 2.0 * SCALE).unwrap(), &up, Transform::identity(), None);
                } else {
                    let mut ip = Paint::default(); ip.set_color_rgba8(theme.inactive_tab_bg.r(), theme.inactive_tab_bg.g(), theme.inactive_tab_bg.b(), theme.inactive_tab_bg.a());
                    pix.fill_rect(Rect::from_xywh(tx_off, top_chrome_px, tw, TAB_BAR_HEIGHT * SCALE).unwrap(), &ip, Transform::identity(), None);
                }

                // Get tab properties for name calculation
                let (is_sticky, name, is_modified) = {
                    let t = &self.tabs[i];
                    (t.is_sticky, t.name.clone(), t.is_modified)
                };
                let name_str = if is_sticky { name } else { format!("{} [P]", name) };
                let col = if active { theme.active_tab_text } else { theme.inactive_tab_text };

                // Tab Text Caching & Rendering
                let tab_mut = &mut self.tabs[i];
                if tab_mut.buffer.is_none() {
                    let mut lab = Buffer::new(&mut self.font_system, Metrics::new(14.0,20.0).scale(SCALE));
                    lab.set_text(&mut self.font_system, &name_str, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None);
                    lab.shape_until_scroll(&mut self.font_system, false);
                    tab_mut.buffer = Some(lab);
                }
                if let Some(lab) = &tab_mut.buffer {
                    for r in lab.layout_runs() {
                        for g in r.glyphs {
                            let pg = g.physical((tx_off + 10.0 * SCALE, top_chrome_px + 10.0 * SCALE + r.line_y), 1.0);
                            if let Some(im) = self.swash_cache.get_image(&mut self.font_system, pg.cache_key) {
                                let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap();
                                let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a());
                                for (idx, &al) in im.data.iter().enumerate() {
                                    let af = (al as f32 / 255.0) * (ca as f32 / 255.0);
                                    p.pixels_mut()[idx] = ColorU8::from_rgba((cr as f32 * af) as u8, (cg as f32 * af) as u8, (cb as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                                }
                                pix.draw_pixmap(pg.x + im.placement.left, pg.y - im.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                            }
                        }
                    }
                }
                
                // Tab close button [X] or Modified dot [•]
                if is_modified {
                    App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "•", tx_off + tw - 24.0 * SCALE, top_chrome_px + 10.0 * SCALE, Color::rgb(180, 180, 180));
                } else {
                    let is_close_hover = self.hovering_tab_close == Some(i);
                    App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "×", tx_off + tw - 24.0 * SCALE, top_chrome_px + 10.0 * SCALE, if is_close_hover { Color::rgb(255, 100, 100) } else { Color::rgb(120,120,120) });
                }

                tx_off += tw;
            }

        // 2b. Breadcrumbs Removed (Unified in Status Bar)

        // 3. Active Editor or Designer
        if self.active_tab < self.tabs.len() {
             let ed_top = top_chrome_px + (TAB_BAR_HEIGHT + UI_BAR_HEIGHT) * SCALE;
             let rect = Rect::from_xywh(ed_start_x, ed_top, pix.width() as f32 - ed_start_x, pix.height() as f32 - (ed_top + FOOTER_HEIGHT * SCALE)).unwrap();

             // Drain LSP events first
             while let Ok(evt) = self.lsp.rx.try_recv() {
                match evt {
                    LspEvent::Diagnostics(uri, diags) => { 
                        for t in &mut self.tabs {
                            let t_uri = t.path.clone().unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", t.name));
                            if t_uri == uri {
                                if let TabContent::Code(cw) = &mut t.content {
                                    // Convert LSP diagnostics to widget's generic DiagnosticInfo
                                    cw.my_editor.diagnostics = diags.iter().map(|d| {
                                        vybe_widgets::DiagnosticInfo {
                                            line: d.range.start.line as usize,
                                            col_start: d.range.start.character as usize,
                                            col_end: d.range.end.character as usize,
                                            message: d.message.clone(),
                                            severity: match d.severity {
                                                Some(lsp_types::DiagnosticSeverity::ERROR) => vybe_widgets::DiagnosticSeverity::Error,
                                                Some(lsp_types::DiagnosticSeverity::WARNING) => vybe_widgets::DiagnosticSeverity::Warning,
                                                Some(lsp_types::DiagnosticSeverity::INFORMATION) => vybe_widgets::DiagnosticSeverity::Info,
                                                _ => vybe_widgets::DiagnosticSeverity::Hint,
                                            },
                                        }
                                    }).collect();
                                    self.needs_redraw = true;
                                    break;
                                }
                            }
                        }
                    }
                    _ => {}
                }
             }

             let tab = &mut self.tabs[self.active_tab];
             match &mut tab.content {
                 TabContent::Form(f) => {
                     f.render(pix, &mut self.font_system, &mut self.swash_cache, crate::form_designer_tab::Rect { x: rect.left() / SCALE, y: rect.top() / SCALE, w: rect.width() / SCALE, h: rect.height() / SCALE }, SCALE);
                 }
                 TabContent::Code(w) => {
                     // Update Editor Wrapping if changed
                     let wrap_lines = w.wrap_lines;
                     w.editor.with_buffer_mut(|b| {
                         let wrap = if wrap_lines { cosmic_text::Wrap::Word } else { cosmic_text::Wrap::None };
                         if b.wrap() != wrap {
                             b.set_wrap(&mut self.font_system, wrap);
                         }
                         if wrap_lines {
                             b.set_size(&mut self.font_system, Some(rect.width() - (GUTTER_WIDTH + MINIMAP_WIDTH) * SCALE), Some(rect.height()));
                         } else {
                             b.set_size(&mut self.font_system, Some(999999.0), Some(999999.0));
                         }
                     });
                     w.needs_reshape = true;
                     w.render(pix, &mut self.font_system, &mut self.swash_cache, rect);
                 }
                 TabContent::Resources(r) => {
                     r.render(pix, &mut self.font_system, &mut self.swash_cache, rect.left() / SCALE, rect.top() / SCALE, rect.width() / SCALE, rect.height() / SCALE, SCALE);
                 }
             }
        }

        // 4. Footer (Enhanced)
        let mut fp = Paint::default(); fp.set_color_rgba8(theme.footer_bg.r(), theme.footer_bg.g(), theme.footer_bg.b(), theme.footer_bg.a());
        pix.fill_rect(Rect::from_xywh(0.0, pix.height() as f32 - FOOTER_HEIGHT * SCALE, pix.width() as f32, FOOTER_HEIGHT * SCALE).unwrap(), &fp, Transform::identity(), None);
        
        if let Some(tab) = self.tabs.get(self.active_tab) {
            let path_str = tab.path.clone().unwrap_or_else(|| tab.name.clone());
            let segments: Vec<&str> = path_str.split(|c| c == '/' || c == '\\').filter(|s| !s.is_empty()).collect();
            
            let status_prefix = match &tab.content {
                TabContent::Code(cw) => {
                    let cursor = cw.editor.cursor();
                    let text = cw.my_editor.rope.to_string();
                    let line_endings = if text.contains("\r\n") { "CRLF" } else { "LF" };
                    let zoom_pct = (cw.font_size / 14.0 * 100.0) as i32;
                    format!("Ln {}, Col {} | {}% | {} | UTF-8 | ", cursor.line + 1, cursor.index + 1, zoom_pct, line_endings)
                }
                TabContent::Form(f) => {
                    let sels = f.selected_controls.len();
                    format!("{} Selected | Form Designer | ", sels)
                }
                TabContent::Resources(r) => {
                    format!("{} resources | {} | ", r.entries.len(), r.active_tab.label())
                }
            };
            
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &status_prefix, 10.0 * SCALE, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
            
            // Draw interactive breadcrumbs
            let mut current_x = 10.0 * SCALE + (status_prefix.len() as f32 * 8.4 * SCALE); // hardcoded approx char width
            self.breadcrumb_rects.clear();
            for (i, seg) in segments.iter().enumerate() {
                let seg_text = if i == segments.len() - 1 { seg.to_string() } else { format!("{} > ", seg) };
                let seg_width = seg_text.len() as f32 * 8.4 * SCALE;
                let rect = Rect::from_xywh(current_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE, seg_width, FOOTER_HEIGHT * SCALE).unwrap();
                
                // Construct full path up to this segment
                let partial_path = segments[0..=i].join("/");
                self.breadcrumb_rects.push((rect, partial_path));
                
                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &seg_text, current_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
                current_x += seg_width;
            }

            let lang_label = format!("Language: {}", self.current_lang);
            let theme_label = format!("Theme: {}", theme_name);
            
            let label_x = pix.width() as f32 - (lang_label.len() as f32 * 9.0 + 20.0) * SCALE;
            let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0) * SCALE;

            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &lang_label, label_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &theme_label, theme_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);

            if let Some(dropdown) = &self.lang_dropdown {
                let (w, h) = dropdown.get_size();
                let menu_x = (pix.width() as f32 / SCALE - w - 20.0).max(10.0);
                let menu_y = (pix.height() as f32 / SCALE - FOOTER_HEIGHT - h - 10.0).max(10.0);
                dropdown.render(
                    pix, &mut self.font_system, &mut self.swash_cache, menu_x, menu_y,
                    (theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), 255),
                    (theme.gutter_divider.r(), theme.gutter_divider.g(), theme.gutter_divider.b(), 255),
                    (theme.selection.r(), theme.selection.g(), theme.selection.b(), 100),
                    (theme.current_line.r(), theme.current_line.g(), theme.current_line.b(), 255),
                    theme.active_tab_text,
                    theme.inactive_tab_text
                );
            }
            if let Some(dropdown) = &self.theme_dropdown {
                let (w, h) = dropdown.get_size();
                let menu_x = (theme_x / SCALE - 10.0).max(10.0);
                // Edge Clamping
                let menu_x = menu_x.min(pix.width() as f32 / SCALE - w - 10.0).max(10.0);
                let menu_y = (pix.height() as f32 / SCALE - FOOTER_HEIGHT - h - 10.0).max(10.0);
                dropdown.render(
                    pix, &mut self.font_system, &mut self.swash_cache, menu_x, menu_y,
                    (theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), 255),
                    (theme.gutter_divider.r(), theme.gutter_divider.g(), theme.gutter_divider.b(), 255),
                    (theme.selection.r(), theme.selection.g(), theme.selection.b(), 100),
                    (theme.current_line.r(), theme.current_line.g(), theme.current_line.b(), 255),
                    theme.active_tab_text,
                    theme.inactive_tab_text
                );
            }
        }

        // 5. Diagnostic Tooltip on Hover
        if let Some(tab) = self.tabs.get(self.active_tab) {
            let _mx = self.mouse_pos.0; let _my = self.mouse_pos.1;
            // Diagnostic tooltips logic...
        }

        // Quick Open Overlay
        if self.is_quick_open {
            let mut o_p = Paint::default(); o_p.set_color_rgba8(30, 30, 35, 240);
            let o_w = 400.0 * SCALE;
            let o_h = 300.0 * SCALE;
            let o_x = (pix.width() as f32 - o_w) / 2.0;
            let o_y = 100.0 * SCALE;
            pix.fill_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap(), &o_p, Transform::identity(), None);
            
            let mut b_p = Paint::default(); b_p.set_color_rgba8(80, 80, 90, 255);
            let mut pb = PathBuilder::new(); pb.push_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap());
            if let Some(path) = pb.finish() { pix.stroke_path(&path, &b_p, &Stroke { width: 1.0 * SCALE, ..Default::default() }, Transform::identity(), None); }
            
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("Go to file: {}|", self.quick_open_query), o_x + 10.0 * SCALE, o_y + 10.0 * SCALE, Color::rgb(200, 200, 200));
            
            let matcher = SkimMatcherV2::default();
            let mut matches: Vec<(i64, usize, &String)> = self.tabs.iter().enumerate()
                .filter_map(|(idx, tab)| {
                    if self.quick_open_query.is_empty() {
                        Some((0, idx, &tab.name))
                    } else {
                        matcher.fuzzy_match(&tab.name, &self.quick_open_query).map(|score| (score, idx, &tab.name))
                    }
                })
                .collect();
            matches.sort_by_key(|m| -m.0); // Highest score first

            let mut i_y = o_y + 50.0 * SCALE;
            for (idx, (score, _tab_idx, name)) in matches.iter().take(10).enumerate() {
                let col = if idx == 0 { Color::rgb(0, 122, 204) } else { Color::rgb(200, 200, 200) };
                let display_text = if *score > 0 { format!("{} (score: {})", name, score) } else { name.to_string() };
                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &display_text, o_x + 20.0 * SCALE, i_y, col);
                i_y += 25.0 * SCALE;
            }
        }

        // Menu dropdown overlay (drawn last, on top of everything)
        if has_form_tab {
            if let Some(form_tab) = self.tabs.iter().find(|t| matches!(&t.content, TabContent::Form(_))) {
                if let TabContent::Form(f) = &form_tab.content {
                    let menu_rect = crate::form_designer_tab::Rect { x: 0.0, y: 0.0, w: pix.width() as f32 / SCALE, h: 28.0 };
                    f.menu_bar.render_dropdown_overlay(pix, &mut self.font_system, &mut self.swash_cache, menu_rect, SCALE);
                }
            }
        }

        // Project properties dialog (modal overlay)
        {
            let win_w = pix.width() as f32 / SCALE;
            let win_h = pix.height() as f32 / SCALE;
            self.project_props_dialog.render(pix, &mut self.font_system, &mut self.swash_cache, win_w, win_h, SCALE, &self.project);
        }

        // Project explorer context menu overlay
        if let Some((cmx, cmy, ref item_name)) = self.pe_context_menu {
            let mut cmp = Paint::default();
            let menu_w = 160.0f32;
            let menu_h = 28.0f32;
            // Shadow
            cmp.set_color_rgba8(0, 0, 0, 40);
            if let Some(r) = Rect::from_xywh((cmx + 2.0) * SCALE, (cmy + 2.0) * SCALE, menu_w * SCALE, menu_h * SCALE) { pix.fill_rect(r, &cmp, Transform::identity(), None); }
            // Background
            cmp.set_color_rgba8(255, 255, 255, 255);
            if let Some(r) = Rect::from_xywh(cmx * SCALE, cmy * SCALE, menu_w * SCALE, menu_h * SCALE) { pix.fill_rect(r, &cmp, Transform::identity(), None); }
            // Border
            cmp.set_color_rgba8(200, 200, 200, 255);
            let mut pb = PathBuilder::new();
            if let Some(r) = Rect::from_xywh(cmx * SCALE, cmy * SCALE, menu_w * SCALE, menu_h * SCALE) { pb.push_rect(r); }
            if let Some(path) = pb.finish() { let mut st = Stroke::default(); st.width = SCALE; pix.stroke_path(&path, &cmp, &st, Transform::identity(), None); }
            // Text
            let label = format!("\u{1F5D1} Remove \"{}\"", item_name);
            crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &label, cmx + 10.0, cmy + 6.0, 12.0, Color::rgba(180, 40, 40, 255), SCALE);
        }

        let mut buffer = surf.buffer_mut().unwrap();
        for (i, p) in pix.pixels().iter().enumerate() {
            buffer[i] = (p.red() as u32) << 16 | (p.green() as u32) << 8 | (p.blue() as u32);
        }
        buffer.present().unwrap();
        self.needs_redraw = false;
    }
    /// Flush code tabs back to project (form code-behind and standalone code files)
    fn flush_code_to_project(&mut self) {
        for tab in &self.tabs {
            if let TabContent::Code(cw) = &tab.content {
                let code = cw.my_editor.rope.to_string();
                let tab_name = &tab.name;
                // Check if this is a form's code-behind (FormName.vb)
                if let Some(form_name) = tab_name.strip_suffix(".vb") {
                    if let Some(fm) = self.project.forms.iter_mut().find(|fm| fm.form.name == form_name) {
                        fm.set_user_code(code);
                        continue;
                    }
                }
                // Otherwise check standalone code files
                if let Some(cf) = self.project.code_files.iter_mut().find(|cf| cf.name == *tab_name) {
                    cf.code = code;
                }
            }
            // Also sync form designer back
            if let TabContent::Form(f) = &tab.content {
                if let Some(fm) = self.project.forms.iter_mut().find(|fm| fm.form.name == f.form.name) {
                    fm.form = f.form.clone();
                }
            }
        }
    }

    /// Save project — flush tabs then open save dialog if no path, else save to existing path
    fn save_project(&mut self) {
        self.flush_code_to_project();
        if let Some(ref path) = self.project_path {
            match vybe_project::serialization::save_project_auto(&self.project, path) {
                Ok(_) => println!("Saved: {}", path),
                Err(e) => println!("Save error: {}", e),
            }
        } else {
            if let Some(path) = rfd::FileDialog::new()
                .add_filter("VB Project", &["vbproj"])
                .save_file()
            {
                let path_str = path.to_string_lossy().to_string();
                self.project_path = Some(path_str.clone());
                match vybe_project::serialization::save_project_auto(&self.project, &path_str) {
                    Ok(_) => println!("Saved: {}", path_str),
                    Err(e) => println!("Save error: {}", e),
                }
            }
        }
    }

    /// Run the project by shelling to vybec
    fn run_project(&mut self) {
        self.flush_code_to_project();
        // Save first if we have a path
        if let Some(ref path) = self.project_path {
            let _ = vybe_project::serialization::save_project_auto(&self.project, path);
            let vybec = std::env::current_exe().ok()
                .and_then(|p| p.parent().map(|d| d.join("vybec")))
                .unwrap_or_else(|| std::path::PathBuf::from("vybec"));
            match std::process::Command::new(&vybec).arg(path).spawn() {
                Ok(child) => {
                    self.run_child = Some(child);
                    println!("Running project: {}", path);
                }
                Err(e) => println!("Could not launch vybec: {}", e),
            }
        } else {
            println!("Save the project first.");
        }
    }

    /// Stop the running project
    fn stop_project(&mut self) {
        if let Some(ref mut child) = self.run_child {
            let _ = child.kill();
            println!("Stopped.");
        }
        self.run_child = None;
    }

    /// Add existing form(s) from disk via file picker
    fn add_existing_form(&mut self) {
        let Some(paths) = rfd::FileDialog::new()
            .set_title("Add Existing Form")
            .add_filter("VB Forms", &["vb"])
            .pick_files()
        else { return };
        for path in paths {
            match vybe_project::load_form_vb(&path) {
                Ok(fm) => {
                    let name = fm.form.name.clone();
                    if self.project.forms.iter().all(|f| f.form.name != name) {
                        let form_clone = fm.form.clone();
                        self.project.forms.push(fm);
                        // Switch designer to this form
                        if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                            if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                                fd.form = form_clone;
                                fd.selected_controls.clear();
                            }
                            self.active_tab = idx;
                        }
                    }
                }
                Err(e) => println!("Failed to load form: {}", e),
            }
        }
    }

    /// Add existing code file(s) from disk via file picker
    fn add_existing_code(&mut self) {
        let Some(paths) = rfd::FileDialog::new()
            .set_title("Add Existing Code File")
            .add_filter("Code Files", &["vb", "bas"])
            .pick_files()
        else { return };
        for path in paths {
            let name = path.file_name().unwrap_or_default().to_string_lossy().to_string();
            let code = match vybe_project::read_text_file(&path) {
                Ok(c) => c,
                Err(e) => { println!("Failed to read: {}", e); continue; }
            };
            if self.project.code_files.iter().all(|cf| cf.name != name) {
                self.project.code_files.push(vybe_project::project::CodeFile {
                    name: name.clone(),
                    code: code.clone(),
                });
                // Open a code tab for it
                let lang = load_language("vb").or_else(|| load_language("rust")).expect("language not found");
                let my_editor = MyEditor::from_text(&code, &lang);
                let uri = format!("file:///project/{}", name);
                let widget = {
                    let text = my_editor.rope.to_string();
                    self.lsp.send(LspRequest::Init(text, "vb".to_string(), uri));
                    CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
                };
                self.tabs.push(Tab { name: name, path: None, content: TabContent::Code(widget), is_sticky: true, buffer: None, is_modified: false });
                self.active_tab = self.tabs.len() - 1;
            }
        }
    }

    /// Remove a project item (form or code file) by name
    fn remove_project_item(&mut self, name: &str) {
        let removed = self.project.remove_form(name) || self.project.remove_code_file(name);
        if removed {
            // Remove any matching tabs
            let tab_name_vb = format!("{}.vb", name);
            self.tabs.retain(|t| t.name != name && t.name != tab_name_vb);
            if self.active_tab >= self.tabs.len() && !self.tabs.is_empty() {
                self.active_tab = self.tabs.len() - 1;
            }
            // If the removed form was showing in the designer, load another form
            if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                    if fd.form.name == name {
                        if let Some(fm) = self.project.forms.first() {
                            fd.form = fm.form.clone();
                            fd.selected_controls.clear();
                        }
                    }
                }
            }
        }
    }

    /// Dispatch an edit action (Undo/Redo/Cut/Copy/Paste/Delete) to the active tab
    fn dispatch_edit_action(&mut self, action: EditAction) {
        if self.active_tab >= self.tabs.len() { return; }
        match &mut self.tabs[self.active_tab].content {
            TabContent::Form(f) => {
                match action {
                    EditAction::Delete => {
                        let sel = f.selected_controls.clone();
                        f.form.controls.retain(|c| !sel.contains(&c.id));
                        f.selected_controls.clear();
                    }
                    EditAction::Cut => {
                        self.control_clipboard = f.selected_controls.iter()
                            .filter_map(|id| f.form.controls.iter().find(|c| c.id == *id).cloned())
                            .collect();
                        let sel = f.selected_controls.clone();
                        f.form.controls.retain(|c| !sel.contains(&c.id));
                        f.selected_controls.clear();
                    }
                    EditAction::Copy => {
                        self.control_clipboard = f.selected_controls.iter()
                            .filter_map(|id| f.form.controls.iter().find(|c| c.id == *id).cloned())
                            .collect();
                    }
                    EditAction::Paste => {
                        let mut new_ids = Vec::new();
                        for orig in &self.control_clipboard {
                            let mut ctrl = orig.clone();
                            ctrl.id = uuid::Uuid::new_v4();
                            ctrl.bounds.x += 20;
                            ctrl.bounds.y += 20;
                            let base = format!("{:?}", ctrl.control_type);
                            let mut max = 0u32;
                            for c in &f.form.controls {
                                if c.name.starts_with(&base) {
                                    if let Ok(n) = c.name[base.len()..].parse::<u32>() { max = max.max(n); }
                                }
                            }
                            ctrl.name = format!("{}{}", base, max + 1);
                            new_ids.push(ctrl.id);
                            f.form.controls.push(ctrl);
                        }
                        f.selected_controls = new_ids;
                    }
                    _ => {} // Undo/Redo not yet supported for form designer
                }
            }
            TabContent::Code(cw) => {
                match action {
                    EditAction::Undo => {
                        let (cl, ci) = { let c = cw.editor.cursor(); (c.line, c.index) };
                        if let Some((text, line, col)) = cw.my_editor.undo(cl, ci) {
                            cw.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                            cw.editor.set_cursor(Cursor::new(line, col));
                        }
                    }
                    EditAction::Redo => {
                        let (cl, ci) = { let c = cw.editor.cursor(); (c.line, c.index) };
                        if let Some((text, line, col)) = cw.my_editor.redo(cl, ci) {
                            cw.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                            cw.editor.set_cursor(Cursor::new(line, col));
                        }
                    }
                    EditAction::Cut => {
                        if let Some(t) = cw.editor.copy_selection() {
                            cw.my_editor.save_snapshot(cw.editor.cursor().line, cw.editor.cursor().index);
                            if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); }
                            cw.editor.action(&mut self.font_system, Action::Delete);
                        }
                    }
                    EditAction::Copy => {
                        if let Some(t) = cw.editor.copy_selection() {
                            if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); }
                        }
                    }
                    EditAction::Paste => {
                        if let Some(cb) = &mut self.clipboard {
                            if let Ok(t) = cb.get_text() {
                                cw.my_editor.save_snapshot(cw.editor.cursor().line, cw.editor.cursor().index);
                                let byte_off = cw.editor.with_buffer(|b| {
                                    let cli = cw.editor.cursor().line;
                                    let mut total = 0;
                                    for i in 0..cli { total += b.lines[i].text().len() + 1; }
                                    total + cw.editor.cursor().index
                                });
                                let (new_line, new_col) = cw.my_editor.insert_string(byte_off, &t, &cw.lang_def);
                                cw.editor.with_buffer_mut(|b| {
                                    b.set_text(&mut self.font_system, &cw.my_editor.rope().to_string(), &Attrs::new().family(Family::Monospace), Shaping::Advanced, None);
                                });
                                cw.editor.set_cursor(Cursor::new(new_line, new_col));
                            }
                        }
                    }
                    EditAction::Delete => {
                        cw.my_editor.save_snapshot(cw.editor.cursor().line, cw.editor.cursor().index);
                        cw.editor.action(&mut self.font_system, Action::Delete);
                    }
                }
                cw.needs_reshape = true;
                cw.sync();
            }
            TabContent::Resources(_) => {}
        }
    }

    /// Create a ResourceEditor pre-populated from project resource files
    fn create_resource_editor_from_project(project: &vybe_project::project::Project) -> vybe_widgets::ResourceEditor {
        let mut editor = vybe_widgets::ResourceEditor::new();
        for rm in &project.resource_files {
            for item in &rm.resources {
                editor.entries.push(vybe_widgets::ResourceEntry {
                    name: item.name.clone(),
                    value: item.value.clone(),
                    comment: item.comment.clone().unwrap_or_default(),
                    tab: match item.resource_type {
                        vybe_project::resources::ResourceType::String => vybe_widgets::ResourceTab::Strings,
                        vybe_project::resources::ResourceType::Image => vybe_widgets::ResourceTab::Images,
                        vybe_project::resources::ResourceType::Icon => vybe_widgets::ResourceTab::Icons,
                        vybe_project::resources::ResourceType::Audio => vybe_widgets::ResourceTab::Audio,
                        vybe_project::resources::ResourceType::File => vybe_widgets::ResourceTab::Files,
                        vybe_project::resources::ResourceType::Other => vybe_widgets::ResourceTab::Other,
                    },
                    file_name: item.file_name.clone(),
                });
            }
        }
        editor
    }

    /// Process a resource editor event — handle add/delete/browse/commit, syncing to project
    fn process_resource_event(evt: vybe_widgets::ResourceEditorEvent, r: &mut vybe_widgets::ResourceEditor, project: &mut vybe_project::project::Project) {
        match evt {
            vybe_widgets::ResourceEditorEvent::AddResource(tab) => {
                if tab.is_file_based() {
                    // Open file picker for file-based resources
                    let mut dialog = rfd::FileDialog::new();
                    let (filter_name, exts): (&str, Vec<&str>) = match tab {
                        vybe_widgets::ResourceTab::Images => ("Images", vec!["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"]),
                        vybe_widgets::ResourceTab::Icons => ("Icons", vec!["ico"]),
                        vybe_widgets::ResourceTab::Audio => ("Audio", vec!["wav", "mp3", "ogg", "flac", "aiff"]),
                        _ => ("All Files", vec!["*"]),
                    };
                    if !exts.is_empty() && exts[0] != "*" {
                        dialog = dialog.add_filter(filter_name, &exts);
                    }
                    dialog = dialog.add_filter("All Files", &["*"]);
                    if let Some(paths) = dialog.pick_files() {
                        for p in paths {
                            let path_str = p.to_string_lossy().to_string();
                            let name = p.file_stem()
                                .and_then(|s| s.to_str())
                                .unwrap_or("Resource1")
                                .replace(|c: char| !c.is_alphanumeric() && c != '_', "_");
                            let res_tab = tab;
                            r.entries.push(vybe_widgets::ResourceEntry {
                                name: name.clone(),
                                value: path_str.clone(),
                                comment: String::new(),
                                tab: res_tab,
                                file_name: Some(path_str.clone()),
                            });
                            // Sync to project
                            let rt = match res_tab {
                                vybe_widgets::ResourceTab::Images => vybe_project::resources::ResourceType::Image,
                                vybe_widgets::ResourceTab::Icons => vybe_project::resources::ResourceType::Icon,
                                vybe_widgets::ResourceTab::Audio => vybe_project::resources::ResourceType::Audio,
                                vybe_widgets::ResourceTab::Files => vybe_project::resources::ResourceType::File,
                                _ => vybe_project::resources::ResourceType::String,
                            };
                            if project.resource_files.is_empty() {
                                project.resource_files.push(vybe_project::ResourceManager::new());
                            }
                            if let Some(rm) = project.resource_files.first_mut() {
                                rm.resources.push(vybe_project::ResourceItem::new_file(name, &path_str, rt));
                            }
                        }
                    }
                } else {
                    // For string/other, just add a blank entry
                    r.entries.push(vybe_widgets::ResourceEntry {
                        name: format!("NewResource{}", r.entries.len() + 1),
                        value: String::new(),
                        comment: String::new(),
                        tab,
                        file_name: None,
                    });
                }
                r.dirty = true;
            }
            vybe_widgets::ResourceEditorEvent::DeleteResource(idx) => {
                if idx < r.entries.len() {
                    r.entries.remove(idx);
                    r.selected_row = None;
                    r.dirty = true;
                    // Sync: rebuild project resources
                    Self::sync_resources_to_project(r, project);
                }
            }
            vybe_widgets::ResourceEditorEvent::BrowseFile(idx) => {
                if idx < r.entries.len() {
                    let entry = &r.entries[idx];
                    let rt = &entry.tab;
                    let mut dialog = rfd::FileDialog::new();
                    let exts: Vec<&str> = match rt {
                        vybe_widgets::ResourceTab::Images => vec!["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"],
                        vybe_widgets::ResourceTab::Icons => vec!["ico"],
                        vybe_widgets::ResourceTab::Audio => vec!["wav", "mp3", "ogg", "flac", "aiff"],
                        _ => vec![],
                    };
                    if !exts.is_empty() {
                        dialog = dialog.add_filter("Supported", &exts);
                    }
                    dialog = dialog.add_filter("All Files", &["*"]);
                    if let Some(path) = dialog.pick_file() {
                        let path_str = path.to_string_lossy().to_string();
                        let name = path.file_stem()
                            .and_then(|s| s.to_str())
                            .unwrap_or("Resource1")
                            .replace(|c: char| !c.is_alphanumeric() && c != '_', "_");
                        r.entries[idx].value = path_str.clone();
                        r.entries[idx].name = name;
                        r.entries[idx].file_name = Some(path_str);
                        r.dirty = true;
                        Self::sync_resources_to_project(r, project);
                    }
                }
            }
            vybe_widgets::ResourceEditorEvent::AddStringResource(name, value, comment) => {
                let tab = r.active_tab;
                r.entries.push(vybe_widgets::ResourceEntry {
                    name: name.clone(),
                    value: value.clone(),
                    comment: comment.clone(),
                    tab,
                    file_name: None,
                });
                r.dirty = true;
                // Sync to project
                if project.resource_files.is_empty() {
                    project.resource_files.push(vybe_project::ResourceManager::new());
                }
                if let Some(rm) = project.resource_files.first_mut() {
                    let mut item = vybe_project::ResourceItem::new_string(name, value);
                    let rt = match tab {
                        vybe_widgets::ResourceTab::Other => vybe_project::resources::ResourceType::Other,
                        _ => vybe_project::resources::ResourceType::String,
                    };
                    item.resource_type = rt;
                    item.comment = if comment.is_empty() { None } else { Some(comment) };
                    rm.resources.push(item);
                }
            }
            vybe_widgets::ResourceEditorEvent::EditCommitted(_, _, _) => {
                Self::sync_resources_to_project(r, project);
            }
            _ => {}
        }
    }

    /// Rebuild project resource_files from resource editor entries
    fn sync_resources_to_project(r: &vybe_widgets::ResourceEditor, project: &mut vybe_project::project::Project) {
        if project.resource_files.is_empty() {
            project.resource_files.push(vybe_project::ResourceManager::new());
        }
        if let Some(rm) = project.resource_files.first_mut() {
            rm.resources.clear();
            for entry in &r.entries {
                let rt = match entry.tab {
                    vybe_widgets::ResourceTab::Strings => vybe_project::resources::ResourceType::String,
                    vybe_widgets::ResourceTab::Images => vybe_project::resources::ResourceType::Image,
                    vybe_widgets::ResourceTab::Icons => vybe_project::resources::ResourceType::Icon,
                    vybe_widgets::ResourceTab::Audio => vybe_project::resources::ResourceType::Audio,
                    vybe_widgets::ResourceTab::Files => vybe_project::resources::ResourceType::File,
                    vybe_widgets::ResourceTab::Other => vybe_project::resources::ResourceType::Other,
                };
                if entry.tab.is_file_based() {
                    let mut item = vybe_project::ResourceItem::new_file(entry.name.clone(), &entry.value, rt);
                    item.comment = if entry.comment.is_empty() { None } else { Some(entry.comment.clone()) };
                    rm.resources.push(item);
                } else {
                    let mut item = vybe_project::ResourceItem::new_string(entry.name.clone(), entry.value.clone());
                    item.resource_type = rt;
                    item.comment = if entry.comment.is_empty() { None } else { Some(entry.comment.clone()) };
                    rm.resources.push(item);
                }
            }
        }
    }

    fn draw_ui_text(pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, text: &str, x: f32, y: f32, col: Color) {
        let mut lab = Buffer::new(fs, Metrics::new(14.0,20.0).scale(SCALE)); lab.set_text(fs, text, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None); lab.shape_until_scroll(fs, false);
        for r in lab.layout_runs() { for g in r.glyphs { let pg = g.physical((x, y + r.line_y), 1.0); if let Some(im) = sc.get_image(fs, pg.cache_key) { let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap(); let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a()); for (idx, &al) in im.data.iter().enumerate() { let af = (al as f32 / 255.0) * (ca as f32 / 255.0); p.pixels_mut()[idx] = ColorU8::from_rgba((cr as f32 * af) as u8, (cg as f32 * af) as u8, (cb as f32 * af) as u8, (255.0 * af) as u8).premultiply(); } pix.draw_pixmap(pg.x + im.placement.left, pg.y - im.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None); } } }
    }
}

impl ApplicationHandler for App {
    fn resumed(&mut self, event_loop: &winit::event_loop::ActiveEventLoop) {
        let window = Arc::new(event_loop.create_window(WindowAttributes::default().with_title("Vybe IDE").with_inner_size(winit::dpi::LogicalSize::new(1200.0, 900.0))).unwrap());
        let ctx = Context::new(window.clone()).unwrap(); let surf = Surface::new(&ctx, window.clone()).unwrap(); let sz = window.inner_size();
        let lang = load_language("rust").expect("load rust");
        let my_editor = MyEditor::from_text("// Welcome to Vybe IDE\nfn main() {\n    println!(\"Multi-file support active!\");\n}", &lang);
        let uri = "file:///Users/youness/www/html/vybe/welcome.rs".to_string();
        let widget = {
                                                let text = my_editor.rope.to_string();
                                                self.lsp.send(LspRequest::Init(text, "rust".to_string(), uri));
                                                CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
                                             };
        self.tabs.push(Tab { name: "welcome.rs".to_string(), path: None, content: TabContent::Code(widget), is_sticky: true, buffer: None, is_modified: false });

        // Always add the Form Designer tab — sync with project's first form
        {
            let mut designer_state = crate::form_designer_tab::FormDesignerState::new();
            if let Some(fm) = self.project.forms.first() {
                designer_state.form = fm.form.clone();
            }
            self.tabs.push(Tab { name: "Form Designer".to_string(), path: None, content: TabContent::Form(designer_state), is_sticky: true, buffer: None, is_modified: false });
        }
        self.active_tab = self.tabs.len().saturating_sub(1);
        self.window = Some(window); self.context = Some(ctx); self.surface = Some(surf);
        self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap()); self.surface.as_mut().unwrap().resize(NonZeroU32::new(sz.width).unwrap(), NonZeroU32::new(sz.height).unwrap()).unwrap();
    }
    fn window_event(&mut self, event_loop: &winit::event_loop::ActiveEventLoop, _id: winit::window::WindowId, event: WindowEvent) {
        match event {
            WindowEvent::CloseRequested => event_loop.exit(), WindowEvent::ModifiersChanged(m) => self.modifiers = m,
            WindowEvent::Resized(sz) => { if let (Some(s), Some(w)) = (&mut self.surface, &self.window) { if sz.width > 0 && sz.height > 0 { s.resize(NonZeroU32::new(sz.width).unwrap(), NonZeroU32::new(sz.height).unwrap()).expect("resize surface"); self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap()); w.request_redraw(); } } }
            WindowEvent::MouseWheel { delta, .. } => { 
                let a = match delta { 
                    MouseScrollDelta::LineDelta(_, y) => y * 120.0, 
                    MouseScrollDelta::PixelDelta(pos) => pos.y as f32 * 2.0 
                }; 
                let tch = self.top_chrome_h();
                if self.mouse_pos.1 / SCALE < tch + TAB_BAR_HEIGHT && self.mouse_pos.1 / SCALE >= tch {
                    self.tab_scroll_x -= a;
                } else if self.active_tab < self.tabs.len() {
                    match &mut self.tabs[self.active_tab].content {
                        TabContent::Code(w) => w.scroll_y -= a,
                        TabContent::Form(f) => {
                            let ed_sx = self.explorer_width + SPLITTER_WIDTH + 1.0;
                            let ed_top = tch + TAB_BAR_HEIGHT;
                            let ph = self.pixmap.as_ref().map(|p| p.height() as f32).unwrap_or(800.0) / SCALE;
                            let pw = self.pixmap.as_ref().map(|p| p.width() as f32).unwrap_or(0.0) / SCALE;
                            let form_rect = crate::form_designer_tab::Rect { x: ed_sx, y: ed_top, w: pw - ed_sx, h: ph - ed_top - FOOTER_HEIGHT };
                            let lay = f.layout(form_rect);
                            let lmx = self.mouse_pos.0 / SCALE;
                            let lmy = self.mouse_pos.1 / SCALE;
                            if lay.toolbox.contains(lmx, lmy) {
                                f.toolbox.scroll(a, lay.toolbox);
                            } else if lay.properties.contains(lmx, lmy) {
                                f.scroll_properties(a);
                            } else if lay.content.contains(lmx, lmy) {
                                f.scroll_y -= a;
                            }
                        }
                        TabContent::Resources(r) => {
                            let ph = self.pixmap.as_ref().map(|p| p.height() as f32).unwrap_or(800.0) / SCALE;
                            r.scroll(a, ph - tch - TAB_BAR_HEIGHT - FOOTER_HEIGHT);
                        }
                    }
                }
                self.window.as_ref().unwrap().request_redraw(); 
            }
            WindowEvent::KeyboardInput { event, .. } => {
                if event.state == ElementState::Pressed {
                    if self.tabs.is_empty() { return; }
                    // Handle Form tab keyboard events
                    if let TabContent::Form(f) = &mut self.tabs[self.active_tab].content {
                        let cmd = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                        let key = &event.logical_key;
                        match key {
                            Key::Named(NamedKey::Delete) | Key::Named(NamedKey::Backspace) => {
                                let sel = f.selected_controls.clone();
                                f.form.controls.retain(|c| !sel.contains(&c.id));
                                f.selected_controls.clear();
                            }
                            Key::Named(NamedKey::Escape) => {
                                if self.project_props_dialog.visible {
                                    self.project_props_dialog.close();
                                } else {
                                    f.selected_controls.clear();
                                    f.menu_bar.open_menu = None;
                                }
                            }
                            Key::Character(c) if cmd => {
                                let ch = c.to_lowercase();
                                match ch.as_str() {
                                    "a" => {
                                        // Select all visual controls
                                        f.selected_controls = f.form.controls.iter()
                                            .filter(|c| !c.control_type.is_non_visual())
                                            .map(|c| c.id).collect();
                                    }
                                    "c" => {
                                        // Copy selected controls
                                        self.control_clipboard = f.selected_controls.iter()
                                            .filter_map(|id| f.form.controls.iter().find(|c| c.id == *id).cloned())
                                            .collect();
                                    }
                                    "x" => {
                                        // Cut selected controls
                                        self.control_clipboard = f.selected_controls.iter()
                                            .filter_map(|id| f.form.controls.iter().find(|c| c.id == *id).cloned())
                                            .collect();
                                        let sel = f.selected_controls.clone();
                                        f.form.controls.retain(|c| !sel.contains(&c.id));
                                        f.selected_controls.clear();
                                    }
                                    "v" => {
                                        // Paste controls with new IDs and offset
                                        let mut new_ids = Vec::new();
                                        for orig in &self.control_clipboard {
                                            let mut ctrl = orig.clone();
                                            ctrl.id = uuid::Uuid::new_v4();
                                            ctrl.bounds.x += 20;
                                            ctrl.bounds.y += 20;
                                            // Auto-rename to avoid duplicates
                                            let base = format!("{:?}", ctrl.control_type);
                                            let mut max = 0u32;
                                            for c in &f.form.controls {
                                                if c.name.starts_with(&base) {
                                                    if let Ok(n) = c.name[base.len()..].parse::<u32>() { max = max.max(n); }
                                                }
                                            }
                                            ctrl.name = format!("{}{}", base, max + 1);
                                            new_ids.push(ctrl.id);
                                            f.form.controls.push(ctrl);
                                        }
                                        f.selected_controls = new_ids;
                                    }
                                    "s" => {
                                        self.save_project();
                                    }
                                    "n" => {
                                        // New project
                                        self.project = vybe_project::project::Project::new("Project1".to_string());
                                        let mut form = vybe_forms::Form::new("Form1".to_string());
                                        form.width = 640; form.height = 480;
                                        self.project.forms.push(vybe_project::project::FormModule::new_classic(form.clone()));
                                        f.form = form;
                                        f.selected_controls.clear();
                                    }
                                    "o" => {
                                        // Open project
                                        if let Some(path) = rfd::FileDialog::new()
                                            .add_filter("VB Project", &["vbproj", "vbp"])
                                            .pick_file()
                                        {
                                            let path_str = path.to_string_lossy().to_string();
                                            if let Ok(proj) = vybe_project::serialization::load_project_auto(&path_str) {
                                                if let Some(first) = proj.forms.first() {
                                                    f.form = first.form.clone();
                                                    f.selected_controls.clear();
                                                }
                                                self.project = proj;
                                                self.project_path = Some(path_str);
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            _ => {}
                        }
                        self.window.as_ref().unwrap().request_redraw();
                        return;
                    }

                    // Handle Resources tab keyboard events
                    if let TabContent::Resources(r) = &mut self.tabs[self.active_tab].content {
                        let key = &event.logical_key;
                        let cmd = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                        match key {
                            Key::Named(NamedKey::Escape) => {
                                if self.project_props_dialog.visible {
                                    self.project_props_dialog.close();
                                } else {
                                    r.handle_escape();
                                }
                            }
                            Key::Named(NamedKey::Enter) => {
                                let evt = r.handle_enter();
                                Self::process_resource_event(evt, r, &mut self.project);
                            }
                            Key::Named(NamedKey::Tab) => {
                                let evt = r.handle_tab();
                                Self::process_resource_event(evt, r, &mut self.project);
                            }
                            Key::Named(NamedKey::Delete) => {
                                let evt = r.handle_delete();
                                Self::process_resource_event(evt, r, &mut self.project);
                            }
                            Key::Named(NamedKey::Backspace) => {
                                r.handle_key('\x08');
                            }
                            Key::Named(NamedKey::ArrowLeft) => {
                                r.handle_left();
                            }
                            Key::Named(NamedKey::ArrowRight) => {
                                r.handle_right();
                            }
                            Key::Named(NamedKey::Home) => {
                                r.handle_home();
                            }
                            Key::Named(NamedKey::End) => {
                                r.handle_end();
                            }
                            _ => {
                                if let Some(t) = &event.text {
                                    if !cmd {
                                        for ch in t.chars() {
                                            if !ch.is_control() {
                                                r.handle_key(ch);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        self.window.as_ref().unwrap().request_redraw();
                        return;
                    }

                    let mut acted = true;
                    let tab = &mut self.tabs[self.active_tab];
                    let w = match &mut tab.content {
                        TabContent::Code(cw) => cw,
                        _ => { self.window.as_ref().unwrap().request_redraw(); return; }
                    };
                    let _theme = w.theme.clone();
                    let cmd = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                    let alt = self.modifiers.state().alt_key(); let shift = self.modifiers.state().shift_key();

                    let key_str = match event.key_without_modifiers() {
                        Key::Character(c) => c.to_lowercase(),
                        Key::Named(nk) => format!("{:?}", nk),
                        _ => String::new(),
                    };

                    for kb in &self.keybindings {
                        if kb.key == key_str && kb.cmd == cmd && kb.shift == shift && kb.alt == alt {
                            match kb.action.as_str() {
                                "Undo" => {
                                    let (cl, ci) = { let c = w.editor.cursor(); (c.line, c.index) };
                                    if let Some((text, line, col)) = w.my_editor.undo(cl, ci) {
                                        w.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                                        w.editor.set_cursor(Cursor::new(line, col));
                                    }
                                }
                                "Redo" => {
                                    let (cl, ci) = { let c = w.editor.cursor(); (c.line, c.index) };
                                    if let Some((text, line, col)) = w.my_editor.redo(cl, ci) {
                                        w.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                                        w.editor.set_cursor(Cursor::new(line, col));
                                    }
                                }
                                "SelectAll" => {
                                    w.editor.set_cursor(Cursor::new(0, 0));
                                    let last_line = w.editor.with_buffer(|b| b.lines.len() - 1);
                                    let last_col = w.editor.with_buffer(|b| b.lines[last_line].text().len());
                                    w.editor.set_selection(Selection::Normal(Cursor::new(last_line, last_col)));
                                }
                                "MoveBufferStart" => w.editor.action(&mut self.font_system, Action::Motion(Motion::BufferStart)),
                                "MoveBufferEnd" => w.editor.action(&mut self.font_system, Action::Motion(Motion::BufferEnd)),
                                "MoveLineStart" => w.editor.action(&mut self.font_system, Action::Motion(Motion::Home)),
                                "MoveLineEnd" => w.editor.action(&mut self.font_system, Action::Motion(Motion::End)),
                                "MoveWordLeft" => w.editor.action(&mut self.font_system, Action::Motion(Motion::LeftWord)),
                                "MoveWordRight" => w.editor.action(&mut self.font_system, Action::Motion(Motion::RightWord)),
                                "Save" => { println!("Saving document: {}", tab.name); tab.is_modified = false; }
                                "Find" => { w.is_search_open = true; if let Some(t) = w.editor.copy_selection() { if !t.is_empty() { w.search_query = t; } } else { w.search_query.clear(); } }
                                "Replace" => { w.is_search_open = true; w.is_replace_open = !w.is_replace_open; }
                                _ => { acted = false; }
                            }
                            if acted { w.needs_reshape = true; w.sync(); self.window.as_ref().unwrap().request_redraw(); return; }
                        }
                    }

                    match event.key_without_modifiers() {
                        Key::Character(c) if cmd && (c == "=" || c == "+") => { w.set_zoom(&mut self.font_system, 1.0); }
                        Key::Character(c) if cmd && c == "-" => { w.set_zoom(&mut self.font_system, -1.0); }
                        Key::Character(c) if cmd && c == "0" => { w.font_size = 14.0; w.set_zoom(&mut self.font_system, 0.0); }
                        Key::Character(c) if cmd && (c == "w" || c == "W") => { w.show_whitespace = !w.show_whitespace; }
                        Key::Character(c) if alt && (c == "z" || c == "Z") => { w.wrap_lines = !w.wrap_lines; w.needs_reshape = true; }
                        Key::Character(c) if cmd && (c == "m" || c == "M") => { if let Some(p) = w.my_editor.find_matching_bracket(w.editor.cursor().line, w.editor.cursor().index, &w.lang_def) { w.editor.set_cursor(Cursor::new(p.0, p.1)); } }
                        Key::Character(c) if cmd && (c == "p" || c == "P") => { self.is_quick_open = !self.is_quick_open; self.quick_open_query.clear(); }
                         Key::Named(NamedKey::Home) => {
                              let cli = w.editor.cursor().line; let cur = w.editor.cursor().index;
                              let line_text = w.editor.with_buffer(|b| b.lines[cli].text().to_string());
                              let first_byte_idx = line_text.char_indices().find(|&(_, c)| !c.is_whitespace()).map(|(i, _)| i).unwrap_or(line_text.len());
                              if cur == first_byte_idx { w.editor.action(&mut self.font_system, Action::Motion(Motion::Home)); }
                              else { w.editor.set_cursor(Cursor::new(cli, first_byte_idx)); }
                         }
                        Key::Named(NamedKey::End) => w.editor.action(&mut self.font_system, Action::Motion(Motion::End)),
                        Key::Character(c) if cmd && shift && (c == "k" || c == "K") => { w.editor.action(&mut self.font_system, Action::Motion(Motion::End)); w.editor.action(&mut self.font_system, Action::Backspace); w.editor.action(&mut self.font_system, Action::Motion(Motion::Home)); let len = w.editor.with_buffer(|b| b.lines[w.editor.cursor().line].text().len()); for _ in 0..len { w.editor.action(&mut self.font_system, Action::Delete); } w.editor.action(&mut self.font_system, Action::Delete); }
                        Key::Named(NamedKey::Backspace) => if w.is_search_open { if w.is_replace_open && alt { w.replace_query.pop(); } else { w.search_query.pop(); } } else { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); w.editor.action(&mut self.font_system, Action::Backspace); tab.is_modified = true; }
                        Key::Named(NamedKey::Delete) => { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); w.editor.action(&mut self.font_system, Action::Delete); tab.is_modified = true; }
                        Key::Named(NamedKey::Enter) => {
                            if self.is_quick_open {
                                self.is_quick_open = false;
                                // In a real IDE, we'd fuzzy search and pick. For win, just close.
                            } else if w.is_search_open { w.find_next(&mut self.font_system); } 
                            else { 
                                w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index);
                                let line_idx = w.editor.cursor().line;
                                let byte_off = w.editor.with_buffer(|b| {
                                    let mut total = 0;
                                    for i in 0..line_idx { total += b.lines[i].text().len() + 1; }
                                    total + w.editor.cursor().index
                                });
                                w.my_editor.insert_newline(byte_off, &w.lang_def);
                                w.editor.action(&mut self.font_system, Action::Enter); // ensure cosmic-text syncs too
                                w.needs_reshape = true; w.sync(); tab.is_modified = true;
                            }
                        }
                        Key::Named(NamedKey::Escape) => { self.is_quick_open = false; w.is_search_open = false; w.context_menu = None; }
                        Key::Character(c) if cmd && (c == "c" || c == "C") => { if let Some(t) = w.editor.copy_selection() { if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } } }
                        Key::Character(c) if cmd && (c == "v" || c == "V") => { 
                             if let Some(cb) = &mut self.clipboard { 
                                 if let Ok(t) = cb.get_text() { 
                                     w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index);
                                     
                                     // Handle selection replacement
                                     if let Some((start, end)) = w.editor.selection_bounds() {
                                         let b = |c: Cursor, ed: &CodeEditorWidget| ed.editor.with_buffer(|buf| {
                                             let mut total = 0;
                                             // cosmic-text 0.18 cursor indices are reliable. 
                                             for i in 0..c.line { total += buf.lines[i].text().len() + 1; }
                                             total + c.index
                                         });
                                         let s_off = b(start, w);
                                         let e_off = b(end, w);
                                         w.my_editor.delete_range(s_off.min(e_off), s_off.max(e_off), &w.lang_def);
                                         w.editor.action(&mut self.font_system, Action::Delete);
                                     }

                                     let byte_off = w.editor.with_buffer(|b| {
                                         let cli = w.editor.cursor().line;
                                         let mut total = 0;
                                         for i in 0..cli { total += b.lines[i].text().len() + 1; }
                                         total + w.editor.cursor().index
                                     });
                                     
                                     let (new_line, new_col) = w.my_editor.insert_string(byte_off, &t, &w.lang_def);
                                     
                                     // Sync cosmic-text
                                     w.editor.with_buffer_mut(|b| {
                                         b.set_text(&mut self.font_system, &w.my_editor.rope().to_string(), &Attrs::new().family(Family::Monospace), Shaping::Advanced, None);
                                     });
                                     w.editor.set_cursor(Cursor::new(new_line, new_col));
                                     tab.is_modified = true; 
                                 } 
                             } 
                        }
                        Key::Character(c) if cmd && (c == "x" || c == "X") => { if let Some(t) = w.editor.copy_selection() { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } w.editor.action(&mut self.font_system, Action::Delete); tab.is_modified = true; } }
                        Key::Character(c) if cmd && (c == "d" || c == "D") => {
                             w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index);
                             if w.editor.selection_bounds().is_none() {
                                 let li = w.editor.cursor().line;
                                 w.my_editor.duplicate_line(li);
                             } else if let Some(_t) = w.editor.copy_selection() {
                                 w.find_next(&mut self.font_system); // Simplified "Next occurrence" logic
                             }
                             w.needs_reshape = true; tab.is_modified = true;
                        }
                        Key::Named(NamedKey::ArrowUp) if alt => { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); let li = w.editor.cursor().line; w.my_editor.move_line_up(li); w.needs_reshape = true; tab.is_modified = true; }
                        Key::Named(NamedKey::ArrowDown) if alt => { w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index); let li = w.editor.cursor().line; if shift { w.my_editor.duplicate_line(li); } else { w.my_editor.move_line_down(li); } w.needs_reshape = true; tab.is_modified = true; }
                        Key::Named(NamedKey::ArrowLeft) => {
                            if shift && w.editor.selection_bounds().is_none() { w.editor.set_selection(Selection::Normal(w.editor.cursor())); }
                            w.editor.action(&mut self.font_system, Action::Motion(Motion::Left));
                            if !shift { w.editor.set_selection(Selection::None); }
                        }
                        Key::Named(NamedKey::ArrowRight) => {
                            if shift && w.editor.selection_bounds().is_none() { w.editor.set_selection(Selection::Normal(w.editor.cursor())); }
                            w.editor.action(&mut self.font_system, Action::Motion(Motion::Right));
                            if !shift { w.editor.set_selection(Selection::None); }
                        }
                        Key::Named(NamedKey::ArrowUp) => {
                            if shift && w.editor.selection_bounds().is_none() { w.editor.set_selection(Selection::Normal(w.editor.cursor())); }
                            w.editor.action(&mut self.font_system, Action::Motion(Motion::Up));
                            if !shift { w.editor.set_selection(Selection::None); }
                        }
                        Key::Named(NamedKey::ArrowDown) => {
                            if shift && w.editor.selection_bounds().is_none() { w.editor.set_selection(Selection::Normal(w.editor.cursor())); }
                            w.editor.action(&mut self.font_system, Action::Motion(Motion::Down));
                            if !shift { w.editor.set_selection(Selection::None); }
                        }
                        Key::Character(c) if cmd && c == "z" => { 
                            if shift {
                                if let Some((text, line, col)) = w.my_editor.redo(w.editor.cursor().line, w.editor.cursor().index) {
                                    w.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                                    let safe_line = line.min(w.editor.with_buffer(|b| b.lines.len().saturating_sub(1)));
                                    let safe_col = w.editor.with_buffer(|b| if safe_line < b.lines.len() { col.min(b.lines[safe_line].text().len()) } else { 0 });
                                    w.editor.set_cursor(Cursor::new(safe_line, safe_col));
                                }
                            } else {
                                if let Some((text, line, col)) = w.my_editor.undo(w.editor.cursor().line, w.editor.cursor().index) {
                                    w.editor.with_buffer_mut(|b| b.set_text(&mut self.font_system, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
                                    let safe_line = line.min(w.editor.with_buffer(|b| b.lines.len().saturating_sub(1)));
                                    let safe_col = w.editor.with_buffer(|b| if safe_line < b.lines.len() { col.min(b.lines[safe_line].text().len()) } else { 0 });
                                    w.editor.set_cursor(Cursor::new(safe_line, safe_col));
                                }
                            }
                        }
                        Key::Character(c) if cmd && c == "a" => {
                            w.editor.set_cursor(Cursor::new(0, 0));
                            let last_line = w.editor.with_buffer(|b| b.lines.len() - 1);
                            let last_col = w.editor.with_buffer(|b| b.lines[last_line].text().len());
                            w.editor.set_selection(Selection::Normal(Cursor::new(last_line, last_col)));
                        }
                        _ => { if let Some(t) = event.text { if !cmd {
                            w.my_editor.save_snapshot(w.editor.cursor().line, w.editor.cursor().index);
                            
                            // Handling selection replacement on type
                            if let Some((start, end)) = w.editor.selection_bounds() {
                                 let b = |c: Cursor, ed: &CodeEditorWidget| ed.editor.with_buffer(|buf| {
                                    let mut total = 0;
                                    for i in 0..c.line { total += buf.lines[i].text().len() + 1; }
                                    total + c.index
                                });
                                let s_off = b(start, w);
                                let e_off = b(end, w);
                                w.my_editor.delete_range(s_off.min(e_off), s_off.max(e_off), &w.lang_def);
                                // cosmic-text Action::Insert handles UI deletion
                            }

                            for ch in t.chars() { if !ch.is_control() || ch == '\t' || ch == '\n' { if w.is_search_open { if w.is_replace_open && alt { w.replace_query.push(ch); } else { w.search_query.pop(); w.search_query.push(ch); } } else { 
                             let mut skip = false; if let Some(cl) = match ch { ')'=>Some(')'),'}'=>Some('}'),']'=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { 
                                 let cli = w.editor.cursor().line; let cur = w.editor.cursor().index; 
                                 let line_text = w.editor.with_buffer(|b| b.lines[cli].text().to_string());
                                 let next_ch = line_text[cur..].chars().next();
                                 if next_ch == Some(cl) { w.editor.action(&mut self.font_system, Action::Motion(Motion::Right)); skip = true; } 
                             }
                            if !skip { w.editor.action(&mut self.font_system, Action::Insert(ch)); tab.is_modified = true; if let Some(cl) = match ch { '('=>Some(')'),'{'=>Some('}'),'['=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { w.editor.action(&mut self.font_system, Action::Insert(cl)); w.editor.action(&mut self.font_system, Action::Motion(Motion::Left)); } }
                        } } } } else { acted = false; } } else { acted = false; } }
                    }
                    if acted { 
                        if let TabContent::Code(w) = &mut self.tabs[self.active_tab].content {
                            w.needs_reshape = true; 
                            w.sync(); 
                            self.pending_lsp_update = true;
                            self.last_lsp_update = Instant::now();
                        }
                        self.window.as_ref().unwrap().request_redraw(); 
                    }
                }
            }
            WindowEvent::CursorMoved { position, .. } => { 
                self.mouse_pos = (position.x as f32, position.y as f32); 
                let mx = self.mouse_pos.0 / SCALE;
                let my = self.mouse_pos.1 / SCALE;
                let height = self.pixmap.as_ref().unwrap().height() as f32 / SCALE;

                // 1. Splitter Hover/Drag
                let split_start = self.explorer_width;
                let split_end = self.explorer_width + SPLITTER_WIDTH;
                let was_hovering_splitter = self.hovering_splitter;
                self.hovering_splitter = mx >= split_start && mx <= split_end;

                if self.is_dragging_splitter {
                    self.explorer_width = (mx - SPLITTER_WIDTH / 2.0).max(50.0).min(600.0);
                }

                // Check if hovering a resource editor column separator
                let mut res_col_hover = false;
                if self.active_tab < self.tabs.len() {
                    if let TabContent::Resources(re) = &self.tabs[self.active_tab].content {
                        let esx = self.explorer_width + SPLITTER_WIDTH + 1.0;
                        let tch_r = self.top_chrome_h();
                        let ed_top_log = tch_r + TAB_BAR_HEIGHT;
                        let rw = (self.pixmap.as_ref().unwrap().width() as f32 / SCALE) - esx;
                        let rh = height - ed_top_log - FOOTER_HEIGHT;
                        res_col_hover = re.is_resizing() || re.is_near_separator(mx, esx, ed_top_log, rw, my, rh);
                    }
                }

                if self.hovering_splitter || self.is_dragging_splitter || res_col_hover {
                    self.window.as_ref().unwrap().set_cursor(winit::window::CursorIcon::ColResize);
                } else {
                    self.window.as_ref().unwrap().set_cursor(winit::window::CursorIcon::Default);
                }

                if let Some(mut dropdown) = self.lang_dropdown.take() {
                    let (w, h) = dropdown.get_size();
                    let menu_x = (self.pixmap.as_ref().unwrap().width() as f32 / SCALE - w - 20.0).max(10.0);
                    let menu_y = (height - FOOTER_HEIGHT - h - 10.0).max(10.0);
                    dropdown.handle_mouse(mx, my, menu_x, menu_y, false);
                    self.lang_dropdown = Some(dropdown);
                }
                
                if let Some(mut dropdown) = self.theme_dropdown.take() {
                    let (w, h) = dropdown.get_size();
                    let theme_label = format!("Theme: {}", self.get_theme_name());
                    let lang_label = format!("Language: {}", self.current_lang);
                    let label_x = (self.pixmap.as_ref().unwrap().width() as f32 / SCALE) - (lang_label.len() as f32 * 9.0 + 20.0);
                    let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0);
                    let menu_x = theme_x.min(self.pixmap.as_ref().unwrap().width() as f32 / SCALE - w - 10.0).max(10.0);
                    let menu_y = (height - FOOTER_HEIGHT - h - 10.0).max(10.0);
                    dropdown.handle_mouse(mx, my, menu_x, menu_y, false);
                    self.theme_dropdown = Some(dropdown);
                }

                if self.lang_dropdown.is_some() || self.theme_dropdown.is_some() {
                    self.window.as_ref().unwrap().request_redraw();
                }

                // 2. Tab Close Hover
                let last_tab_hover = self.hovering_tab_close;
                self.hovering_tab_close = None;
                let ed_start_x = self.explorer_width + SPLITTER_WIDTH + 1.0;
                let tch = self.top_chrome_h();
                if my >= tch && my < tch + TAB_BAR_HEIGHT && mx > ed_start_x {
                    let mut tx = ed_start_x;
                    for i in 0..self.tabs.len() {
                        let tw = 160.0;
                        if mx >= tx + tw - 30.0 && mx <= tx + tw - 5.0 {
                            self.hovering_tab_close = Some(i);
                            break;
                        }
                        tx += tw;
                    }
                }

                // 3. Menu hover for form designer
                if let Some(form_tab) = self.tabs.iter_mut().find(|t| matches!(&t.content, TabContent::Form(_))) {
                    if let TabContent::Form(f) = &mut form_tab.content {
                        let menu_rect = crate::form_designer_tab::Rect { x: 0.0, y: 0.0, w: self.pixmap.as_ref().unwrap().width() as f32 / SCALE, h: 28.0 };
                        f.menu_bar.handle_hover(self.mouse_pos.0 / SCALE, self.mouse_pos.1 / SCALE, menu_rect);
                    }
                }

                // 4. Resource column resize (independent of is_dragging)
                let mut needs_editor_redraw = false;
                if self.active_tab < self.tabs.len() {
                    if let TabContent::Resources(re) = &mut self.tabs[self.active_tab].content {
                        if re.is_resizing() {
                            let ed_top = (tch + TAB_BAR_HEIGHT) * SCALE;
                            let rw = (self.pixmap.as_ref().unwrap().width() as f32 - ed_start_x * SCALE) / SCALE;
                            re.handle_col_resize_move(mx, rw);
                            needs_editor_redraw = true;
                        }
                    }
                }

                // 5. Editor Drag
                if self.is_dragging && !self.is_dragging_splitter && self.active_tab < self.tabs.len() {
                    let ed_top = (tch + TAB_BAR_HEIGHT) * SCALE;
                    let r = Rect::from_xywh(ed_start_x * SCALE, ed_top, self.pixmap.as_ref().unwrap().width() as f32 - ed_start_x * SCALE, self.pixmap.as_ref().unwrap().height() as f32 - (ed_top + FOOTER_HEIGHT * SCALE)).unwrap();
                    match &mut self.tabs[self.active_tab].content {
                        TabContent::Code(cw) => cw.handle_mouse(&mut self.font_system, self.mouse_pos.0, self.mouse_pos.1, r, None, &mut self.clipboard),
                        TabContent::Form(f) => f.handle_mouse_move(self.mouse_pos.0 / SCALE, self.mouse_pos.1 / SCALE, crate::form_designer_tab::Rect { x: ed_start_x, y: ed_top / SCALE, w: r.width() / SCALE, h: r.height() / SCALE }),
                        TabContent::Resources(_) => {} // resize handled above
                    }
                    needs_editor_redraw = true;
                }

                // Form designer menu open needs continuous redraw for hover
                let form_menu_open = self.tabs.iter().any(|t| matches!(&t.content, TabContent::Form(f) if f.menu_bar.open_menu.is_some()));

                // Smart Redraw: Only if interaction state changed or dragging
                if was_hovering_splitter != self.hovering_splitter ||
                   last_tab_hover != self.hovering_tab_close ||
                   self.is_dragging_splitter ||
                   self.lang_dropdown.is_some() ||
                   form_menu_open ||
                   res_col_hover ||
                   needs_editor_redraw {
                    self.window.as_ref().unwrap().request_redraw();
                }
            }
            WindowEvent::Focused(false) | WindowEvent::CursorLeft { .. } => {
                self.is_dragging = false;
                self.is_dragging_splitter = false;
            }
                WindowEvent::MouseInput { state, button, .. } => {
                    let mx = self.mouse_pos.0 / SCALE;
                    let my = self.mouse_pos.1 / SCALE;
                    let pw = self.pixmap.as_ref().unwrap().width() as f32;
                    let ph = self.pixmap.as_ref().unwrap().height() as f32 / SCALE;
                    let height = ph; // restore height alias

                    if state == ElementState::Pressed && button == MouseButton::Right {
                        // Right-click in project explorer sidebar → context menu
                        let tch_r = self.top_chrome_h();
                        if mx < self.explorer_width && my > tch_r && self.sidebar_tab == SidebarTab::Project {
                            // Find which item was right-clicked
                            let pe_y = tch_r + SIDEBAR_TAB_H;
                            let item_h = 24.0f32;
                            let mut iy = pe_y - self.project_explorer.scroll_y;
                            iy += item_h; // project name
                            iy += item_h; // forms header
                            if !self.project_explorer.forms_collapsed {
                                for fm in &self.project.forms {
                                    if my >= iy && my < iy + item_h {
                                        self.pe_context_menu = Some((mx, my, fm.form.name.clone()));
                                        self.window.as_ref().unwrap().request_redraw();
                                        return;
                                    }
                                    iy += item_h;
                                }
                            }
                            if !self.project.code_files.is_empty() {
                                iy += item_h; // code header
                                if !self.project_explorer.code_collapsed {
                                    for cf in &self.project.code_files {
                                        if my >= iy && my < iy + item_h {
                                            self.pe_context_menu = Some((mx, my, cf.name.clone()));
                                            self.window.as_ref().unwrap().request_redraw();
                                            return;
                                        }
                                        iy += item_h;
                                    }
                                }
                            }
                        } else if self.active_tab < self.tabs.len() {
                            if let TabContent::Code(cw) = &mut self.tabs[self.active_tab].content {
                                cw.context_menu = Some(((mx, my), vec!["Cut".into(), "Copy".into(), "Paste".into(), "Go to Def".into()]));
                            }
                        }
                        self.window.as_ref().unwrap().request_redraw();
                        return;
                    }

                    if state == ElementState::Pressed && button == MouseButton::Left {
                        // 0. Context menu click (project explorer)
                        if let Some((cmx, cmy, ref item_name)) = self.pe_context_menu.clone() {
                            let menu_w = 160.0f32;
                            let menu_h = 28.0f32;
                            if mx >= cmx && mx < cmx + menu_w && my >= cmy && my < cmy + menu_h {
                                // Clicked "Remove" — remove the item
                                self.remove_project_item(&item_name);
                            }
                            self.pe_context_menu = None;
                            self.window.as_ref().unwrap().request_redraw();
                            return;
                        }

                        // 0. Project Properties Dialog (modal — must be first)
                        if self.project_props_dialog.visible {
                            let win_w = pw / SCALE;
                            let win_h = height;
                            if self.project_props_dialog.is_ok_clicked(mx, my, win_w, win_h) {
                                self.project_props_dialog.apply(&mut self.project);
                                self.project_props_dialog.close();
                            } else {
                                self.project_props_dialog.handle_click(mx, my, win_w, win_h, &self.project);
                            }
                            self.window.as_ref().unwrap().request_redraw();
                            return;
                        }

                        // 0a. Language Picker Menu Intercept
                        if let Some(mut dropdown) = self.lang_dropdown.take() {
                            let (w, h) = dropdown.get_size();
                            let label_x = (pw / SCALE) - (format!("Language: {}", self.current_lang).len() as f32 * 9.0 + 20.0);
                            let menu_x = label_x.min(pw / SCALE - w - 10.0).max(10.0);
                            let menu_y = (height - FOOTER_HEIGHT - h - 10.0).max(10.0);
                            
                            match dropdown.handle_mouse(mx, my, menu_x, menu_y, true) {
                                DropdownEvent::Selected(idx) => {
                                    if let Some(new_lang) = self.all_languages.get(idx).cloned() {
                                        self.current_lang = new_lang.clone();
                                        let tab = &mut self.tabs[self.active_tab];
                                        let uri = tab.path.clone().unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", tab.name));
                                        if let TabContent::Code(cw) = &mut tab.content {
                                            { cw.set_language(&new_lang); let t = cw.my_editor.rope.to_string(); self.lsp.send(LspRequest::Init(t, new_lang.clone(), uri)); };
                                        }
                                    }
                                    self.lang_dropdown = None;
                                }
                                DropdownEvent::Closed => { self.lang_dropdown = None; }
                                DropdownEvent::None => self.lang_dropdown = Some(dropdown),
                            }
                            self.window.as_ref().unwrap().request_redraw(); return;
                        }

                        // 0b. Theme Picker Menu Intercept
                        if let Some(mut dropdown) = self.theme_dropdown.take() {
                            let (w, h) = dropdown.get_size();
                            let theme_label = format!("Theme: {}", self.get_theme_name());
                            let lang_label = format!("Language: {}", self.current_lang);
                            let label_x = (pw / SCALE) - (lang_label.len() as f32 * 9.0 + 20.0);
                            let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0);
                            let menu_x = theme_x.min(pw / SCALE - w - 10.0).max(10.0);
                            let menu_y = (height - FOOTER_HEIGHT - h - 10.0).max(10.0);
                            
                            match dropdown.handle_mouse(mx, my, menu_x, menu_y, true) {
                                DropdownEvent::Selected(idx) => {
                                    self.current_theme_idx = idx;
                                    let new_theme = self.active_theme();
                                    for tab in &mut self.tabs { 
                                        if let TabContent::Code(cw) = &mut tab.content {
                                            cw.theme = new_theme.clone(); cw.needs_reshape = true;
                                        }
                                    }
                                    self.window.as_ref().unwrap().request_redraw();
                                    return;
                                }
                                DropdownEvent::None => self.theme_dropdown = Some(dropdown),
                                _ => {}
                            }
                            self.window.as_ref().unwrap().request_redraw(); return;
                        }

                        // 1. Minimap Hit-testing (code editor only — form designer handles its own right panel)
                        if mx > pw / SCALE - MINIMAP_WIDTH {
                            if self.active_tab < self.tabs.len() {
                                if let TabContent::Code(cw) = &mut self.tabs[self.active_tab].content {
                                    let mut th = 0.0; cw.editor.with_buffer(|b| { for r in b.layout_runs() { if !cw.is_line_hidden(r.line_i) { th += r.line_height; } } });
                                    let mry = (my - TAB_BAR_HEIGHT) / (height - TAB_BAR_HEIGHT - FOOTER_HEIGHT);
                                    cw.scroll_y = (mry * th).max(0.0);
                                    self.window.as_ref().unwrap().request_redraw();
                                    return;
                                }
                            }
                        }


                        // 3. Status Bar Click
                        if my >= height - FOOTER_HEIGHT {
                            // Breadcrumb segments hit-testing
                            for (rect, path) in &self.breadcrumb_rects {
                                if mx * SCALE >= rect.left() && mx * SCALE <= rect.right() && my * SCALE >= rect.top() && my * SCALE <= rect.bottom() {
                                    println!("Revealing in explorer: {}", path);
                                    continue;
                                }
                            }

                            let lang_label = format!("Language: {}", self.current_lang);
                            let theme_label = format!("Theme: {}", self.get_theme_name());
                            let label_x = (pw / SCALE) - (lang_label.len() as f32 * 9.0 + 20.0);
                            let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0);

                            if mx >= label_x {
                                let active_idx = self.all_languages.iter().position(|l| l == &self.current_lang).unwrap_or(0);
                                self.lang_dropdown = Some(Dropdown::new(self.all_languages.clone(), active_idx, SCALE, None));
                            } else if mx >= theme_x && mx < label_x {
                                let theme_names = vec![
                                    "Silicon Green".into(), "Cloud Blue".into(), "Coffee Cream".into(), "Sakura Pink".into(), 
                                    "One Dark".into(), "Monokai".into(), "GitHub Light".into(), "Solarized Light".into(), 
                                    "Midnight".into(), "Aura".into(), "Veridian".into(), "Rose".into(),
                                    "Cyber".into(), "Titanium".into(), "Indigo Night".into()
                                ];
                                let mut d = Dropdown::new(theme_names, self.current_theme_idx, SCALE, None);
                                d.num_cols = 2; d.col_w = 160.0;
                                self.theme_dropdown = Some(d);
                            }
                            self.window.as_ref().unwrap().request_redraw(); return;
                        }

                        // 4a. Menu bar / toolbar / dropdown click
                    let tch = self.top_chrome_h();
                    if tch > 0.0 {
                        // Check if a menu dropdown is open — must handle dropdown clicks first
                        let menu_open = self.tabs.iter().any(|t| matches!(&t.content, TabContent::Form(ref f) if f.menu_bar.open_menu.is_some()));

                        if my < tch || menu_open {
                            if let Some(form_tab) = self.tabs.iter_mut().find(|t| matches!(&t.content, TabContent::Form(_))) {
                                if let TabContent::Form(f) = &mut form_tab.content {
                                    let menu_rect = crate::form_designer_tab::Rect { x: 0.0, y: 0.0, w: pw / SCALE, h: 28.0 };
                                    if let Some(action) = f.menu_bar.handle_click(mx, my, menu_rect) {
                                        self.window.as_ref().unwrap().request_redraw();
                                        // Dispatch menu action
                                        match action {
                                            crate::form_designer_tab::MenuAction::NewProject => {
                                                self.project = vybe_project::project::Project::new("Project1".to_string());
                                                let mut form = vybe_forms::Form::new("Form1".to_string());
                                                form.width = 640; form.height = 480;
                                                let fm = vybe_project::project::FormModule::new_classic(form);
                                                self.project.forms.push(fm);
                                                // Switch to form designer
                                                if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                                                    self.active_tab = idx;
                                                    if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                                                        fd.form = self.project.forms[0].form.clone();
                                                    }
                                                }
                                            }
                                            crate::form_designer_tab::MenuAction::OpenProject => {
                                                if let Some(path) = rfd::FileDialog::new()
                                                    .add_filter("VB Project", &["vbproj", "vbp"])
                                                    .pick_file()
                                                {
                                                    let path_str = path.to_string_lossy().to_string();
                                                    match vybe_project::serialization::load_project_auto(&path_str) {
                                                        Ok(proj) => {
                                                            // Close any existing code tabs (stale content)
                                                            self.tabs.retain(|t| !matches!(&t.content, TabContent::Code(_)) && !matches!(&t.content, TabContent::Resources(_)));
                                                            if self.active_tab >= self.tabs.len() && !self.tabs.is_empty() {
                                                                self.active_tab = self.tabs.len() - 1;
                                                            }
                                                            if let Some(first_form) = proj.forms.first() {
                                                                if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                                                                    if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                                                                        fd.form = first_form.form.clone();
                                                                        fd.selected_controls.clear();
                                                                    }
                                                                    self.active_tab = idx;
                                                                }
                                                            }
                                                            self.project = proj;
                                                            self.project_path = Some(path_str);
                                                        }
                                                        Err(e) => { println!("Error loading project: {}", e); }
                                                    }
                                                }
                                            }
                                            crate::form_designer_tab::MenuAction::SaveProject => {
                                                self.save_project();
                                            }
                                            crate::form_designer_tab::MenuAction::SaveAs => {
                                                // Force file picker even if we have a path
                                                self.flush_code_to_project();
                                                if let Some(path) = rfd::FileDialog::new()
                                                    .add_filter("VB Project", &["vbproj"])
                                                    .save_file()
                                                {
                                                    let path_str = path.to_string_lossy().to_string();
                                                    self.project_path = Some(path_str.clone());
                                                    match vybe_project::serialization::save_project_auto(&self.project, &path_str) {
                                                        Ok(_) => { println!("Saved: {}", path_str); }
                                                        Err(e) => { println!("Save error: {}", e); }
                                                    }
                                                }
                                            }
                                            crate::form_designer_tab::MenuAction::Exit => {
                                                std::process::exit(0);
                                            }
                                            crate::form_designer_tab::MenuAction::AddForm => {
                                                let name = format!("Form{}", self.project.forms.len() + 1);
                                                let mut form = vybe_forms::Form::new(name.clone());
                                                form.width = 640; form.height = 480;
                                                let form_clone = form.clone();
                                                let fm = vybe_project::project::FormModule::new_classic(form);
                                                self.project.forms.push(fm);
                                                // Switch designer to new form
                                                if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                                                    if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                                                        fd.form = form_clone;
                                                        fd.selected_controls.clear();
                                                    }
                                                    self.active_tab = idx;
                                                }
                                            }
                                            crate::form_designer_tab::MenuAction::AddModule => {
                                                let name = format!("Module{}.vb", self.project.code_files.len() + 1);
                                                self.project.code_files.push(vybe_project::project::CodeFile {
                                                    name: name.clone(),
                                                    code: format!("Module {}\n\nEnd Module\n", name.replace(".vb", "")),
                                                });
                                            }
                                            crate::form_designer_tab::MenuAction::AddResourceFile => {
                                                // Auto-create ResourceManager in project if missing (like legacy_ide)
                                                if self.project.resource_files.is_empty() {
                                                    self.project.resource_files.push(vybe_project::ResourceManager::new());
                                                }
                                                // Open or switch to Resources tab
                                                if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Resources(_))) {
                                                    self.active_tab = idx;
                                                } else {
                                                    let editor = Self::create_resource_editor_from_project(&self.project);
                                                    self.tabs.push(Tab { name: "Resources.resx".to_string(), path: None, content: TabContent::Resources(editor), is_sticky: true, buffer: None, is_modified: false });
                                                    self.active_tab = self.tabs.len() - 1;
                                                }
                                            }
                                            crate::form_designer_tab::MenuAction::ProjectProperties => {
                                                self.project_props_dialog.open(&self.project);
                                            }
                                            crate::form_designer_tab::MenuAction::RunProject => {
                                                self.run_project();
                                            }
                                            crate::form_designer_tab::MenuAction::StopProject => {
                                                self.stop_project();
                                            }
                                            crate::form_designer_tab::MenuAction::AddExistingForm => {
                                                self.add_existing_form();
                                            }
                                            crate::form_designer_tab::MenuAction::AddExistingCode => {
                                                self.add_existing_code();
                                            }
                                            crate::form_designer_tab::MenuAction::Undo => {
                                                self.dispatch_edit_action(EditAction::Undo);
                                            }
                                            crate::form_designer_tab::MenuAction::Redo => {
                                                self.dispatch_edit_action(EditAction::Redo);
                                            }
                                            crate::form_designer_tab::MenuAction::Cut => {
                                                self.dispatch_edit_action(EditAction::Cut);
                                            }
                                            crate::form_designer_tab::MenuAction::Copy => {
                                                self.dispatch_edit_action(EditAction::Copy);
                                            }
                                            crate::form_designer_tab::MenuAction::Paste => {
                                                self.dispatch_edit_action(EditAction::Paste);
                                            }
                                            crate::form_designer_tab::MenuAction::Delete => {
                                                self.dispatch_edit_action(EditAction::Delete);
                                            }
                                        }
                                        return;
                                    }
                                    // If no menu action but clicked in toolbar area
                                    if my >= 28.0 && my < tch {
                                        if let Some(action) = crate::form_designer_tab::toolbar_handle_click_pub(mx, my, crate::form_designer_tab::Rect { x: 0.0, y: 28.0, w: pw / SCALE, h: 36.0 }) {
                                            match action {
                                                crate::form_designer_tab::ToolbarAction::Save => {
                                                    self.save_project();
                                                }
                                                crate::form_designer_tab::ToolbarAction::Run => {
                                                    self.run_project();
                                                }
                                                crate::form_designer_tab::ToolbarAction::Stop => {
                                                    self.stop_project();
                                                }
                                                crate::form_designer_tab::ToolbarAction::ViewDesigner => {
                                                    // If currently on a code tab for a form, switch to that form's designer
                                                    let current_form_name = if self.active_tab < self.tabs.len() {
                                                        let tab = &self.tabs[self.active_tab];
                                                        // Code tab named "FormX.vb" → form name is "FormX"
                                                        if matches!(&tab.content, TabContent::Code(_)) {
                                                            tab.name.strip_suffix(".vb").map(|s| s.to_string())
                                                        } else { None }
                                                    } else { None };

                                                    if let Some(ref fname) = current_form_name {
                                                        // Load this form into the designer
                                                        if let Some(fm) = self.project.forms.iter().find(|fm| &fm.form.name == fname) {
                                                            let form_clone = fm.form.clone();
                                                            if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                                                                if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                                                                    fd.form = form_clone;
                                                                    fd.selected_controls.clear();
                                                                }
                                                                self.active_tab = idx;
                                                            }
                                                        }
                                                    } else {
                                                        // Just switch to existing Form tab
                                                        if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                                                            self.active_tab = idx;
                                                        }
                                                    }
                                                }
                                                crate::form_designer_tab::ToolbarAction::ViewCode => {
                                                    // Get the current form name from the active form designer tab
                                                    let form_name = if self.active_tab < self.tabs.len() {
                                                        if let TabContent::Form(f) = &self.tabs[self.active_tab].content {
                                                            Some(f.form.name.clone())
                                                        } else { None }
                                                    } else { None };
                                                    let form_name = form_name.unwrap_or_else(|| {
                                                        self.tabs.iter().find_map(|t| {
                                                            if let TabContent::Form(f) = &t.content { Some(f.form.name.clone()) } else { None }
                                                        }).unwrap_or_else(|| "Module1".to_string())
                                                    });

                                                    let code_tab_name = format!("{}.vb", form_name);

                                                    // Find existing code tab for this form
                                                    if let Some(idx) = self.tabs.iter().position(|t| t.name == code_tab_name && matches!(&t.content, TabContent::Code(_))) {
                                                        self.active_tab = idx;
                                                    } else {
                                                        // Get code-behind from FormModule::get_user_code() — NOT code_files
                                                        let code = self.project.forms.iter()
                                                            .find(|fm| fm.form.name == form_name)
                                                            .map(|fm| fm.get_user_code().to_string())
                                                            .unwrap_or_else(|| format!("Public Class {}\n\nEnd Class\n", form_name));

                                                        let lang = load_language("vb").or_else(|| load_language("rust")).expect("language not found");
                                                        let my_editor = MyEditor::from_text(&code, &lang);
                                                        let uri = format!("file:///project/{}", code_tab_name);
                                                        let widget = {
                                                            let text = my_editor.rope.to_string();
                                                            self.lsp.send(LspRequest::Init(text, "vb".to_string(), uri));
                                                            CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
                                                        };
                                                        self.tabs.push(Tab { name: code_tab_name, path: None, content: TabContent::Code(widget), is_sticky: true, buffer: None, is_modified: false });
                                                        self.active_tab = self.tabs.len() - 1;
                                                    }
                                                }
                                                crate::form_designer_tab::ToolbarAction::AddForm => {
                                                    let name = format!("Form{}", self.project.forms.len() + 1);
                                                    let mut form = vybe_forms::Form::new(name.clone());
                                                    form.width = 640; form.height = 480;
                                                    let form_clone = form.clone();
                                                    self.project.forms.push(vybe_project::project::FormModule::new_classic(form));
                                                    // Switch designer to new form
                                                    if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                                                        if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                                                            fd.form = form_clone;
                                                            fd.selected_controls.clear();
                                                        }
                                                        self.active_tab = idx;
                                                    }
                                                }
                                                crate::form_designer_tab::ToolbarAction::AddCode => {
                                                    let name = format!("Module{}.vb", self.project.code_files.len() + 1);
                                                    let code = format!("Module {}\n\nEnd Module\n", name.replace(".vb", ""));
                                                    self.project.code_files.push(vybe_project::project::CodeFile {
                                                        name: name.clone(),
                                                        code: code.clone(),
                                                    });
                                                    // Also open a code tab for it
                                                    let lang = load_language("vb").or_else(|| load_language("rust")).expect("language not found");
                                                    let my_editor = MyEditor::from_text(&code, &lang);
                                                    let uri = format!("file:///project/{}", name);
                                                    let widget = {
                                                        let text = my_editor.rope.to_string();
                                                        self.lsp.send(LspRequest::Init(text, "vb".to_string(), uri));
                                                        CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
                                                    };
                                                    self.tabs.push(Tab { name: name, path: None, content: TabContent::Code(widget), is_sticky: true, buffer: None, is_modified: false });
                                                    self.active_tab = self.tabs.len() - 1;
                                                }
                                            }
                                            self.window.as_ref().unwrap().request_redraw();
                                            return;
                                        }
                                    }
                                    // If dropdown was open but clicked outside, it closed — absorb click
                                    if menu_open {
                                        self.window.as_ref().unwrap().request_redraw(); return;
                                    }
                                }
                            }
                            if my < tch {
                                self.window.as_ref().unwrap().request_redraw(); return;
                            }
                        }
                    }

                    // 4. Tab Bar Click
                    let ed_start_x = self.explorer_width + SPLITTER_WIDTH + 1.0;
                    if my >= tch && my < tch + TAB_BAR_HEIGHT && mx > ed_start_x {
                        if let Some(idx) = self.hovering_tab_close {
                             println!("DEBUG: Tab Bar Click -> Close Tab {}", idx);
                             self.tabs.remove(idx);
                             if self.tabs.is_empty() { self.active_tab = 0; }
                             else if self.active_tab >= self.tabs.len() { self.active_tab = self.tabs.len() - 1; }
                             self.window.as_ref().unwrap().request_redraw(); return;
                        }
                        let tab_idx = ((mx - ed_start_x) / 160.0) as usize;
                        if tab_idx < self.tabs.len() { 
                            println!("DEBUG: Tab Bar Click -> Select Tab {}", tab_idx);
                            self.active_tab = tab_idx; self.window.as_ref().unwrap().request_redraw(); 
                        }
                        return;
                    }

                    // 3. Sidebar Click
                    if mx < self.explorer_width {
                        // Sidebar tab click
                        let stab_top = tch;
                        if my >= stab_top && my < stab_top + SIDEBAR_TAB_H {
                            let half = self.explorer_width / 2.0;
                            if mx < half {
                                self.sidebar_tab = SidebarTab::Files;
                            } else {
                                self.sidebar_tab = SidebarTab::Project;
                            }
                            self.window.as_ref().unwrap().request_redraw(); return;
                        }

                        let now = Instant::now();
                        let is_double = (now - self.last_click_time) < Duration::from_millis(300);
                        self.last_click_time = now;

                        match self.sidebar_tab {
                            SidebarTab::Files => {
                                match self.tree_view.handle_mouse(self.mouse_pos.0, self.mouse_pos.1, 0.0, (tch + SIDEBAR_TAB_H) * SCALE) {
                                     TreeEvent::Open(path) => {
                                         if let Some(idx) = self.tabs.iter().position(|t| t.path.as_ref() == Some(&path)) {
                                             self.active_tab = idx; if is_double { self.tabs[idx].is_sticky = true; }
                                             self.window.as_ref().unwrap().request_redraw(); return;
                                         }
                                         if let Ok(content) = fs::read_to_string(&path) {
                                             let ext = path.split('.').last().unwrap_or("txt");
                                             let lang_name = match ext { "rs" => "rust", "js" => "javascript", "bas" | "vb" => "vb", "cs" => "csharp", _ => "text" };
                                             let lang = load_language(lang_name).or_else(|| load_language("rust")).expect("rust language not found");
                                             let my_editor = MyEditor::from_text(&content, &lang);
                                             let uri = format!("file://{}", path);
                                             let mut widget = {
                                                let text = my_editor.rope.to_string();
                                                self.lsp.send(LspRequest::Init(text, "rust".to_string(), uri.clone()));
                                                CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
                                             };
                                             { widget.set_language(lang_name); let t = widget.my_editor.rope.to_string(); self.lsp.send(LspRequest::Init(t, lang_name.to_string(), uri)); };
                                             let name = Path::new(&path).file_name().unwrap_or_default().to_string_lossy().to_string();
                                             let new_tab = Tab { name, path: Some(path.clone()), content: TabContent::Code(widget), is_sticky: is_double, buffer: None, is_modified: false };
                                             self.tabs.push(new_tab); self.active_tab = self.tabs.len() - 1; self.tree_view.reveal_path(&path);
                                         }
                                     }
                                     _ => {}
                                }
                            }
                            SidebarTab::Project => {
                                // Project explorer click — handle collapse toggles
                                let pe_y = tch + SIDEBAR_TAB_H;
                                let item_h = 24.0f32;
                                let mut iy = pe_y - self.project_explorer.scroll_y;

                                // Project name — skip
                                iy += item_h;

                                // Forms header
                                if my >= iy && my < iy + item_h {
                                    self.project_explorer.forms_collapsed = !self.project_explorer.forms_collapsed;
                                    self.window.as_ref().unwrap().request_redraw(); return;
                                }
                                iy += item_h;

                                if !self.project_explorer.forms_collapsed {
                                    for i in 0..self.project.forms.len() {
                                        if my >= iy && my < iy + item_h {
                                            // Load this form into the Form Designer tab
                                            let form_clone = self.project.forms[i].form.clone();
                                            if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                                                if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                                                    fd.form = form_clone;
                                                    fd.selected_controls.clear();
                                                }
                                                self.active_tab = idx;
                                            }
                                            self.window.as_ref().unwrap().request_redraw(); return;
                                        }
                                        iy += item_h;
                                    }
                                }

                                // Code header
                                if !self.project.code_files.is_empty() {
                                    if my >= iy && my < iy + item_h {
                                        self.project_explorer.code_collapsed = !self.project_explorer.code_collapsed;
                                        self.window.as_ref().unwrap().request_redraw(); return;
                                    }
                                    iy += item_h;
                                    if !self.project_explorer.code_collapsed {
                                        for i in 0..self.project.code_files.len() {
                                            if my >= iy && my < iy + item_h {
                                                // Open code file in a code tab
                                                let cf = &self.project.code_files[i];
                                                let lang = load_language("vb").or_else(|| load_language("rust")).expect("language not found");
                                                let my_editor = MyEditor::from_text(&cf.code, &lang);
                                                let uri = format!("file:///project/{}", cf.name);
                                                let widget = {
                                                let text = my_editor.rope.to_string();
                                                self.lsp.send(LspRequest::Init(text, "rust".to_string(), uri));
                                                CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
                                             };
                                                let new_tab = Tab { name: cf.name.clone(), path: None, content: TabContent::Code(widget), is_sticky: true, buffer: None, is_modified: false };
                                                self.tabs.push(new_tab);
                                                self.active_tab = self.tabs.len() - 1;
                                                self.window.as_ref().unwrap().request_redraw(); return;
                                            }
                                            iy += item_h;
                                        }
                                    }
                                }

                                // References header
                                if !self.project.project_references.is_empty() {
                                    if my >= iy && my < iy + item_h {
                                        self.project_explorer.refs_collapsed = !self.project_explorer.refs_collapsed;
                                        self.window.as_ref().unwrap().request_redraw(); return;
                                    }
                                    iy += item_h;
                                    if !self.project_explorer.refs_collapsed {
                                        for _ in &self.project.project_references {
                                            iy += item_h;
                                        }
                                    }
                                }

                                // Resources header — only if resources exist or Resources tab is open
                                let has_any_res = (!self.project.resource_files.is_empty() &&
                                    self.project.resource_files.iter().any(|rm| !rm.resources.is_empty() || rm.file_path.is_some()))
                                    || self.tabs.iter().any(|t| matches!(&t.content, TabContent::Resources(_)));
                                if has_any_res {
                                if my >= iy && my < iy + item_h {
                                    self.project_explorer.resources_collapsed = !self.project_explorer.resources_collapsed;
                                    self.window.as_ref().unwrap().request_redraw(); return;
                                }
                                iy += item_h;
                                if !self.project_explorer.resources_collapsed {
                                    for _ in 0..self.project.resource_files.len() {
                                        if my >= iy && my < iy + item_h {
                                            // Open/switch to Resources tab
                                            if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Resources(_))) {
                                                self.active_tab = idx;
                                            } else {
                                                let editor = Self::create_resource_editor_from_project(&self.project);
                                                self.tabs.push(Tab { name: "Resources.resx".to_string(), path: None, content: TabContent::Resources(editor), is_sticky: true, buffer: None, is_modified: false });
                                                self.active_tab = self.tabs.len() - 1;
                                            }
                                            self.window.as_ref().unwrap().request_redraw(); return;
                                        }
                                        iy += item_h;
                                    }
                                }
                                } // end if has_any_res
                            }
                        }
                        self.window.as_ref().unwrap().request_redraw(); return;
                    }

                    // 4. Splitter Click
                    if self.hovering_splitter {
                        println!("DEBUG: Splitter Click -> Start Resizing");
                        self.is_dragging_splitter = true; self.window.as_ref().unwrap().request_redraw(); return;
                    }

                    // 5. Editor Click (Deep Isolation)
                    if !self.tabs.is_empty() {
                        let ed_top = (tch + TAB_BAR_HEIGHT) * SCALE;
                        let ed_bottom = height * SCALE - FOOTER_HEIGHT * SCALE;
                        if self.mouse_pos.1 >= ed_top && self.mouse_pos.1 < ed_bottom && mx >= self.explorer_width + SPLITTER_WIDTH {
                            println!("DEBUG: Editor Click at mx={}, my={}", mx, my);
                            let rect = Rect::from_xywh(ed_start_x * SCALE, ed_top, self.pixmap.as_ref().unwrap().width() as f32 - ed_start_x * SCALE, ed_bottom - ed_top).unwrap();
                            self.click_count = if Instant::now().duration_since(self.last_click_time) < Duration::from_millis(500) { (self.click_count % 3) + 1 } else { 1 }; self.last_click_time = Instant::now();
                            
                            let mut start_drag = true;
                            match &mut self.tabs[self.active_tab].content {
                                TabContent::Code(cw) => {
                                    cw.handle_mouse(&mut self.font_system, self.mouse_pos.0, self.mouse_pos.1, rect, Some((self.click_count, button, self.modifiers)), &mut self.clipboard);
                                }
                                TabContent::Form(f) => {
                                    let ctrl_held = self.modifiers.state().control_key() || self.modifiers.state().super_key();
                                    let form_rect = crate::form_designer_tab::Rect { x: ed_start_x, y: ed_top / SCALE, w: rect.width() / SCALE, h: rect.height() / SCALE };
                                    let handled = f.handle_mouse_down(self.mouse_pos.0 / SCALE, self.mouse_pos.1 / SCALE, form_rect, ctrl_held);
                                    // Only start drag if the click was on the form canvas content area
                                    let lay = f.layout(form_rect);
                                    let lmx = self.mouse_pos.0 / SCALE;
                                    let lmy = self.mouse_pos.1 / SCALE;
                                    start_drag = handled && lay.content.contains(lmx, lmy);
                                }
                                TabContent::Resources(r) => {
                                    let rx = ed_start_x;
                                    let ry = ed_top / SCALE;
                                    let rw = rect.width() / SCALE;
                                    let rh = rect.height() / SCALE;
                                    // Check column resize first
                                    if r.handle_col_resize_start(mx, my, rx, ry, rw, rh) {
                                        start_drag = true; // use is_dragging to track resize
                                    } else {
                                        let evt = r.handle_click(mx, my, rx, ry, rw, rh);
                                        Self::process_resource_event(evt, r, &mut self.project);
                                        start_drag = false;
                                    }
                                }
                            }
                            if start_drag { self.is_dragging = true; }
                            self.window.as_ref().unwrap().request_redraw();
                        }
                    }
                } else if state == ElementState::Released {
                    println!("DEBUG: Mouse Released");
                    let tch_rel = self.top_chrome_h();
                    if self.active_tab < self.tabs.len() {
                        if let TabContent::Form(f) = &mut self.tabs[self.active_tab].content {
                            let ed_sx = self.explorer_width + SPLITTER_WIDTH + 1.0;
                            let ed_top = tch_rel + TAB_BAR_HEIGHT;
                            let ed_h = height - ed_top - FOOTER_HEIGHT;
                            let ed_w = pw / SCALE - ed_sx;
                            f.handle_mouse_up(crate::form_designer_tab::Rect { x: ed_sx, y: ed_top, w: ed_w, h: ed_h });
                        }
                    }
                    // End resource column resize
                    if self.active_tab < self.tabs.len() {
                        if let TabContent::Resources(r) = &mut self.tabs[self.active_tab].content {
                            r.handle_col_resize_end();
                        }
                    }
                    self.is_dragging = false;
                    self.is_dragging_splitter = false;
                    self.window.as_ref().unwrap().request_redraw();
                }
            }
            WindowEvent::RedrawRequested => self.render(),
            _ => {}
        }
    }
}

pub fn run_gui(my_editor: MyEditor, open_form: bool) {
    let el = EventLoop::new().expect("event loop"); el.set_control_flow(ControlFlow::Wait);
    let mut app = App::new(my_editor, open_form); el.run_app(&mut app).expect("run app");
}
