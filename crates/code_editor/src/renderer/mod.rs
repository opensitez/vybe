mod dialogs;
mod app_render;
mod app_commands;
mod app_keyboard;
mod app_mouse;

use std::time::Instant;
use std::sync::Arc;
use std::fs;
use cosmic_text::{Attrs, Buffer, Color, Family, FontSystem, Metrics, Shaping, SwashCache};
use tiny_skia::{Pixmap, PixmapPaint, Rect, Transform, ColorU8};
use winit::event::ElementState;
use arboard::Clipboard;

use serde::{Deserialize, Serialize};

use crate::editor::Editor as MyEditor;
use crate::language::load_language;
use crate::lsp_client::{LspClient, LspRequest};
use vybe_widgets::{TreeView, Dropdown};
use vybe_widgets::code_editor_widget::{Theme, CodeEditorWidget};

use dialogs::ProjectPropsDialog;

// ── Types ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct Keybinding {
    pub key: String,
    #[serde(default)] pub cmd: bool,
    #[serde(default)] pub shift: bool,
    #[serde(default)] pub alt: bool,
    pub action: String,
}

// ── Layout Constants ───────────────────────────────────────────────────

pub(crate) const SCALE: f32 = 2.0;
pub(crate) const EXPLORER_WIDTH: f32 = 250.0;
pub(crate) const TAB_BAR_HEIGHT: f32 = 36.0;
pub(crate) const MINIMAP_WIDTH: f32 = 80.0;
pub(crate) const UI_BAR_HEIGHT: f32 = 0.0;
pub(crate) const FOOTER_HEIGHT: f32 = 24.0;
pub(crate) const GUTTER_WIDTH: f32 = 64.0;
pub(crate) const SPLITTER_WIDTH: f32 = 4.0;
pub(crate) const SIDEBAR_TAB_H: f32 = 28.0;

// ── Enums ──────────────────────────────────────────────────────────────

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum SidebarTab { Files, Project }

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum BottomPanelTab { Output, Problems }

#[derive(Clone, Copy)]
pub(crate) enum EditAction { Undo, Redo, Cut, Copy, Paste, Delete }

// ── Helper Structs ─────────────────────────────────────────────────────

pub(crate) struct ProjectExplorerState {
    pub(crate) scroll_y: f32,
    pub(crate) forms_collapsed: bool,
    pub(crate) code_collapsed: bool,
    pub(crate) refs_collapsed: bool,
    pub(crate) resources_collapsed: bool,
}

impl ProjectExplorerState {
    fn new() -> Self { Self { scroll_y: 0.0, forms_collapsed: false, code_collapsed: false, refs_collapsed: false, resources_collapsed: false } }
}

// ── Tab Types ──────────────────────────────────────────────────────────

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

// ── App ────────────────────────────────────────────────────────────────

pub(crate) struct App {
    pub(crate) font_system: FontSystem,
    pub(crate) swash_cache: SwashCache,
    pub(crate) win_width: f32,
    pub(crate) win_height: f32,
    pub(crate) cursor: winit::window::CursorIcon,
    pub(crate) tabs: Vec<Tab>,
    pub(crate) active_tab: usize,
    pub(crate) tree_view: TreeView,
    pub(crate) all_languages: Vec<String>,
    pub(crate) current_lang: String,
    pub(crate) lang_dropdown: Option<Dropdown>,
    pub(crate) theme_dropdown: Option<Dropdown>,
    pub(crate) clipboard: Option<Clipboard>,
    pub(crate) cmd_held: bool,
    pub(crate) shift_held: bool,
    pub(crate) alt_held: bool,
    pub(crate) last_click_time: Instant,
    pub(crate) click_count: u32,
    pub(crate) mouse_pos: (f32, f32),
    pub(crate) explorer_width: f32,
    pub(crate) is_dragging_splitter: bool,
    pub(crate) hovering_splitter: bool,
    pub(crate) hovering_tab_close: Option<usize>,
    pub(crate) is_dragging: bool,
    pub(crate) needs_redraw: bool,
    pub(crate) last_lsp_update: Instant,
    pub(crate) pending_lsp_update: bool,
    pub(crate) lsp: Arc<LspClient>,
    pub(crate) is_quick_open: bool,
    pub(crate) quick_open_query: String,
    pub(crate) tab_scroll_x: f32,
    pub(crate) current_theme_idx: usize,
    pub(crate) breadcrumb_rects: Vec<(Rect, String)>,
    pub(crate) keybindings: Vec<Keybinding>,
    pub(crate) _open_form: bool,
    pub(crate) sidebar_tab: SidebarTab,
    pub(crate) project: vybe_project::project::Project,
    pub(crate) project_explorer: ProjectExplorerState,
    pub(crate) project_props_dialog: ProjectPropsDialog,
    pub(crate) control_clipboard: Vec<vybe_forms::Control>,
    pub(crate) project_path: Option<String>,
    pub(crate) run_child: Option<std::process::Child>,
    pub(crate) pe_context_menu: Option<(f32, f32, String)>,
    pub(crate) last_hover_pos: (f32, f32),
    pub(crate) last_hover_time: Instant,
    // Output panel
    pub(crate) output_lines: Vec<String>,
    pub(crate) output_visible: bool,
    pub(crate) output_panel_height: f32,
    pub(crate) output_scroll_y: f32,
    pub(crate) bottom_panel_tab: BottomPanelTab,
    // Go-to-line dialog
    pub(crate) goto_line_open: bool,
    pub(crate) goto_line_query: String,
    // Build configuration
    pub(crate) build_config: BuildConfig,
}

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum BuildConfig { Debug, Release }

impl App {
    fn new(_my_editor: MyEditor, open_form: bool) -> Self {
        let mut langs = Vec::new();
        let mut candidates: Vec<std::path::PathBuf> = vec![
            std::path::PathBuf::from("crates/code_editor/basic-languages"),
            std::path::PathBuf::from("basic-languages"),
            std::path::PathBuf::from("../code_editor/basic-languages"),
        ];
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
            font_system: FontSystem::new(), swash_cache: SwashCache::new(),
            win_width: 0.0, win_height: 0.0, cursor: winit::window::CursorIcon::Default,
            tabs: Vec::new(), active_tab: 0, 
            tree_view: TreeView::new(".", 2.0), all_languages: langs, current_lang: "rust".to_string(), lang_dropdown: None, theme_dropdown: None, clipboard: Clipboard::new().ok(), 
            cmd_held: false, shift_held: false, alt_held: false, last_click_time: Instant::now(), click_count: 0, mouse_pos: (0.0, 0.0), 
            explorer_width: EXPLORER_WIDTH, is_dragging_splitter: false, hovering_splitter: false, hovering_tab_close: None,
            is_dragging: false, needs_redraw: true,
            last_lsp_update: Instant::now(), pending_lsp_update: false,
            lsp: Arc::new(LspClient::new()),
            is_quick_open: false,
            quick_open_query: String::new(),
            tab_scroll_x: 0.0,
            current_theme_idx: 0,
            breadcrumb_rects: Vec::new(),
            _open_form: open_form,
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
            last_hover_pos: (-1.0, -1.0),
            last_hover_time: Instant::now(),
            output_lines: Vec::new(),
            output_visible: false,
            output_panel_height: 150.0,
            output_scroll_y: 0.0,
            bottom_panel_tab: BottomPanelTab::Output,
            goto_line_open: false,
            goto_line_query: String::new(),
            build_config: BuildConfig::Debug,
        }
    }

    /// Height in logical pixels of menu bar + toolbar (always present if any Form tab exists).
    pub(crate) fn top_chrome_h(&self) -> f32 {
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

    fn draw_ui_text(pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, text: &str, x: f32, y: f32, col: Color) {
        let mut lab = Buffer::new(fs, Metrics::new(14.0,20.0).scale(SCALE)); lab.set_text(fs, text, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None); lab.shape_until_scroll(fs, false);
        for r in lab.layout_runs() { for g in r.glyphs { let pg = g.physical((x, y + r.line_y), 1.0); if let Some(im) = sc.get_image(fs, pg.cache_key) { let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap(); let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a()); for (idx, &al) in im.data.iter().enumerate() { let af = (al as f32 / 255.0) * (ca as f32 / 255.0); p.pixels_mut()[idx] = ColorU8::from_rgba((cr as f32 * af) as u8, (cg as f32 * af) as u8, (cb as f32 * af) as u8, (255.0 * af) as u8).premultiply(); } pix.draw_pixmap(pg.x + im.placement.left, pg.y - im.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None); } } }
    }
}

// ── Application (toolkit) ──────────────────────────────────────────────

impl vybe_widgets::Application for App {
    fn on_init(&mut self, width: f32, height: f32, _scale: f32) {
        self.win_width = width;
        self.win_height = height;

        let lang = load_language("rust").expect("load rust");
        let my_editor = MyEditor::from_text("// Welcome to Vybe IDE\nfn main() {\n    println!(\"Multi-file support active!\");\n}", &lang);
        let uri = "file:///Users/youness/www/html/vybe/welcome.rs".to_string();
        let widget = {
            let text = my_editor.rope.to_string();
            self.lsp.send(LspRequest::Init(text, "rust".to_string(), uri));
            CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
        };
        self.tabs.push(Tab { name: "welcome.rs".to_string(), path: None, content: TabContent::Code(widget), is_sticky: true, buffer: None, is_modified: false });

        // Always add the Form Designer tab
        {
            let mut designer_state = crate::form_designer_tab::FormDesignerState::new();
            if let Some(fm) = self.project.forms.first() {
                designer_state.form = fm.form.clone();
            }
            self.tabs.push(Tab { name: "Form Designer".to_string(), path: None, content: TabContent::Form(designer_state), is_sticky: true, buffer: None, is_modified: false });
        }
        self.active_tab = self.tabs.len().saturating_sub(1);
    }

    fn on_resize(&mut self, width: f32, height: f32) {
        self.win_width = width;
        self.win_height = height;
    }

    fn render(&mut self, pix: &mut Pixmap, _scale: f32) {
        self.render_internal(pix);
    }

    fn handle_mouse(&mut self, event: vybe_widgets::MouseEvent) -> bool {
        use vybe_widgets::layout::{MouseEventKind, MouseButton};
        self.cmd_held = event.cmd;
        self.shift_held = event.shift;
        self.alt_held = event.alt;
        self.mouse_pos = (event.x * SCALE, event.y * SCALE);
        match event.kind {
            MouseEventKind::Move => {
                self.handle_cursor_moved();
            }
            MouseEventKind::Press(btn) | MouseEventKind::Release(btn) => {
                let winit_btn = match btn {
                    MouseButton::Left => winit::event::MouseButton::Left,
                    MouseButton::Right => winit::event::MouseButton::Right,
                    MouseButton::Middle => winit::event::MouseButton::Middle,
                };
                let winit_state = if matches!(event.kind, MouseEventKind::Press(_)) {
                    ElementState::Pressed
                } else {
                    ElementState::Released
                };
                self.handle_mouse_input(winit_state, winit_btn);
            }
            _ => {}
        }
        true
    }

    fn handle_key(&mut self, event: vybe_widgets::KeyEvent) -> bool {
        self.cmd_held = event.cmd;
        self.shift_held = event.shift;
        self.alt_held = event.alt;
        self.handle_key_press(event);
        true
    }

    fn handle_scroll(&mut self, delta: f32, _x: f32, _y: f32) -> bool {
        // Convert to winit-style delta and reuse existing handler
        let winit_delta = winit::event::MouseScrollDelta::PixelDelta(
            winit::dpi::PhysicalPosition::new(0.0, delta as f64 / 2.0)
        );
        self.handle_mouse_wheel(winit_delta);
        true
    }

    fn cursor_icon(&self) -> winit::window::CursorIcon {
        self.cursor
    }

    fn on_focus_lost(&mut self) {
        self.is_dragging = false;
        self.is_dragging_splitter = false;
    }
}

// ── Entry Point ────────────────────────────────────────────────────────

pub fn run_gui(my_editor: MyEditor, open_form: bool) {
    let app = App::new(my_editor, open_form);
    vybe_widgets::run_app("Vybe IDE", 1200, 900, SCALE, app);
}
