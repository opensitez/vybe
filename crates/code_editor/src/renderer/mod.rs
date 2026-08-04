mod app_commands;
mod app_keyboard;
mod app_mouse;
mod app_render;
mod dialogs;

use arboard::Clipboard;
use std::fs;
use std::sync::Arc;
use std::time::Instant;
use tiny_skia::{Pixmap, Rect};
use vybe_widgets::{FontSystem, SwashCache, TextColor};
use winit::event::ElementState;

use serde::{Deserialize, Serialize};

use crate::editor::Editor as MyEditor;
use crate::lsp_client::LspClient;
use vybe_widgets::code_editor_widget::{CodeEditorWidget, Theme};
use vybe_widgets::layout::WidgetEvent;
use vybe_widgets::output_panel::{OutputPanel, OutputPanelEvent};
use vybe_widgets::{Dropdown, TreeView};
use vybe_widgets::{LayoutRect, PanelWidget, SplitPanel, StatusBarPanel, TabPanel};

use dialogs::ProjectPropsDialog;

// ── Types ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct Keybinding {
    pub key: String,
    #[serde(default)]
    pub cmd: bool,
    #[serde(default)]
    pub shift: bool,
    #[serde(default)]
    pub alt: bool,
    pub action: String }

// ── Layout Constants ───────────────────────────────────────────────────

pub(crate) const SCALE: f32 = 2.0;
pub(crate) const EXPLORER_WIDTH: f32 = 250.0;
pub(crate) const TAB_BAR_HEIGHT: f32 = 36.0;
pub(crate) const MINIMAP_WIDTH: f32 = 80.0;
pub(crate) const UI_BAR_HEIGHT: f32 = 22.0;
pub(crate) const FOOTER_HEIGHT: f32 = 24.0;
pub(crate) const GUTTER_WIDTH: f32 = 64.0;
pub(crate) const SPLITTER_WIDTH: f32 = 4.0;
pub(crate) const SIDEBAR_TAB_H: f32 = 28.0;

// ── Enums ──────────────────────────────────────────────────────────────

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum SidebarTab {
    Files,
    Project }

#[derive(Clone, Copy, PartialEq)]
#[allow(dead_code)]
pub(crate) enum BottomPanelTab {
    Output,
    Problems }

#[derive(Clone, Copy)]
pub(crate) enum EditAction {
    Undo,
    Redo,
    Cut,
    Copy,
    Paste,
    Delete }

/// A single action runnable from the command palette.
#[derive(Clone)]
pub(crate) struct PaletteCommand {
    pub(crate) label: &'static str,
    pub(crate) action: PaletteAction }

#[derive(Clone, Copy)]
pub(crate) enum PaletteAction {
    Menu(crate::form_designer_tab::MenuAction),
    Edit(EditAction),
    ToggleOutput,
    ToggleProblems,
    CloseTab,
    CloseOthers,
    CloseAll,
    NextTab,
    PrevTab,
    FindInFile,
    FindInProject,
    GoToLine }

pub(crate) fn palette_commands() -> &'static [PaletteCommand] {
    use crate::form_designer_tab::MenuAction::*;
    use PaletteAction::*;
    &[
        PaletteCommand {
            label: "File: New Project",
            action: Menu(NewProject) },
        PaletteCommand {
            label: "File: Open Project…",
            action: Menu(OpenProject) },
        PaletteCommand {
            label: "File: Save Project",
            action: Menu(SaveProject) },
        PaletteCommand {
            label: "File: Save Project As…",
            action: Menu(SaveAs) },
        PaletteCommand {
            label: "File: Exit",
            action: Menu(Exit) },
        PaletteCommand {
            label: "Edit: Undo",
            action: Edit(EditAction::Undo) },
        PaletteCommand {
            label: "Edit: Redo",
            action: Edit(EditAction::Redo) },
        PaletteCommand {
            label: "Edit: Cut",
            action: Edit(EditAction::Cut) },
        PaletteCommand {
            label: "Edit: Copy",
            action: Edit(EditAction::Copy) },
        PaletteCommand {
            label: "Edit: Paste",
            action: Edit(EditAction::Paste) },
        PaletteCommand {
            label: "Edit: Delete",
            action: Edit(EditAction::Delete) },
        PaletteCommand {
            label: "Project: Add Form",
            action: Menu(AddForm) },
        PaletteCommand {
            label: "Project: Add Module",
            action: Menu(AddModule) },
        PaletteCommand {
            label: "Project: Add Existing Form…",
            action: Menu(AddExistingForm) },
        PaletteCommand {
            label: "Project: Add Existing Code…",
            action: Menu(AddExistingCode) },
        PaletteCommand {
            label: "Project: Add Resource File",
            action: Menu(AddResourceFile) },
        PaletteCommand {
            label: "Project: Properties…",
            action: Menu(ProjectProperties) },
        PaletteCommand {
            label: "Run: Start",
            action: Menu(RunProject) },
        PaletteCommand {
            label: "Run: Stop",
            action: Menu(StopProject) },
        PaletteCommand {
            label: "View: Toggle Output Panel",
            action: ToggleOutput },
        PaletteCommand {
            label: "View: Show Problems",
            action: ToggleProblems },
        PaletteCommand {
            label: "View: Close Tab",
            action: CloseTab },
        PaletteCommand {
            label: "View: Close Other Tabs",
            action: CloseOthers },
        PaletteCommand {
            label: "View: Close All Tabs",
            action: CloseAll },
        PaletteCommand {
            label: "View: Next Tab",
            action: NextTab },
        PaletteCommand {
            label: "View: Previous Tab",
            action: PrevTab },
        PaletteCommand {
            label: "Find: In File",
            action: FindInFile },
        PaletteCommand {
            label: "Find: In Project",
            action: FindInProject },
        PaletteCommand {
            label: "Go: To Line",
            action: GoToLine },
    ]
}

/// A hit from a project-wide text search.
#[derive(Clone, Debug)]
pub(crate) struct ProjectSearchHit {
    pub(crate) file: String,
    pub(crate) line: usize,
    pub(crate) snippet: String }

// ── Helper Structs ─────────────────────────────────────────────────────

pub(crate) struct ProjectExplorerState {
    pub(crate) scroll_y: f32,
    pub(crate) forms_collapsed: bool,
    pub(crate) code_collapsed: bool,
    pub(crate) refs_collapsed: bool,
    pub(crate) resources_collapsed: bool }

impl ProjectExplorerState {
    fn new() -> Self {
        Self {
            scroll_y: 0.0,
            forms_collapsed: false,
            code_collapsed: false,
            refs_collapsed: false,
            resources_collapsed: false }
    }
}

// ── Tab Types ──────────────────────────────────────────────────────────

pub enum TabContent {
    Code(CodeEditorWidget),
    Form(crate::form_designer_tab::FormDesignerState),
    Resources(vybe_widgets::ResourceEditor) }

pub struct Tab {
    pub name: String,
    pub path: Option<String>,
    pub content: TabContent,
    pub is_sticky: bool,
    #[allow(dead_code)]
    pub buffer: Option<()>,
    pub is_modified: bool }

// ── App ────────────────────────────────────────────────────────────────

pub(crate) struct App {
    pub(crate) font_system: FontSystem,
    pub(crate) swash_cache: SwashCache,
    pub(crate) win_width: f32,
    pub(crate) win_height: f32,
    /// Actual OS display scale factor (from winit). Used to create dropdowns
    /// so render_list multiplies by the right factor.
    pub(crate) display_scale: f32,
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
    #[allow(dead_code)]
    pub(crate) is_dragging: bool,
    pub(crate) needs_redraw: bool,
    pub(crate) last_lsp_update: Instant,
    pub(crate) pending_lsp_update: bool,
    #[allow(dead_code)]
    pub(crate) lsp: Arc<LspClient>,
    pub(crate) is_quick_open: bool,
    pub(crate) quick_open_query: String,
    pub(crate) is_command_palette: bool,
    pub(crate) command_palette_query: String,
    pub(crate) command_palette_selected: usize,
    pub(crate) is_project_search: bool,
    pub(crate) project_search_query: String,
    pub(crate) project_search_results: Vec<ProjectSearchHit>,
    pub(crate) project_search_selected: usize,
    /// `(screen_x, screen_y, tab_idx)` when a right-click menu is open on a tab.
    pub(crate) tab_context_menu: Option<(f32, f32, usize)>,
    /// Index of the tab currently being dragged (for reorder), if any.
    pub(crate) tab_drag_idx: Option<usize>,
    #[allow(dead_code)]
    pub(crate) tab_scroll_x: f32,
    pub(crate) current_theme_idx: usize,
    pub(crate) breadcrumb_rects: Vec<(Rect, String)>,
    pub(crate) keybindings: Vec<Keybinding>,
    pub(crate) _open_form: bool,
    pub(crate) sidebar_tab: SidebarTab,
    pub(crate) project: vybe_platform_dotnet::winforms::designer::project::Project,
    pub(crate) project_explorer: ProjectExplorerState,
    pub(crate) project_props_dialog: ProjectPropsDialog,
    pub(crate) control_clipboard: Vec<vybe_platform_dotnet::winforms::designer::Control>,
    pub(crate) project_path: Option<String>,
    pub(crate) run_child: Option<std::process::Child>,
    pub(crate) pe_context_menu: Option<(f32, f32, String)>,
    pub(crate) last_hover_pos: (f32, f32),
    pub(crate) last_hover_time: Instant,
    // Output panel
    pub(crate) output_panel: OutputPanel,
    pub(crate) output_panel_height: f32,
    pub(crate) output_lines_buffer: Vec<String>,
    // Go-to-line dialog
    pub(crate) goto_line_open: bool,
    pub(crate) goto_line_query: String,
    // Build configuration
    pub(crate) build_config: BuildConfig,
    // ── Toolkit Containers ──────────────────────────────────────────
    pub(crate) split_panel: SplitPanel,    // sidebar | editor area
    pub(crate) tab_panel: TabPanel,        // editor tab bar
    pub(crate) status_bar: StatusBarPanel, // footer status bar
    pub(crate) sidebar_tabs: TabPanel,     // sidebar Files|Project tabs
}

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum BuildConfig {
    Debug,
    Release }

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
                let cand = dir
                    .join("crates")
                    .join("code_editor")
                    .join("basic-languages");
                if cand.exists() {
                    candidates.insert(0, cand);
                    break;
                }
                dir_opt = dir.parent();
            }
        }
        for path in &candidates {
            if let Ok(es) = std::fs::read_dir(path) {
                for e in es.flatten() {
                    if e.file_type().map(|t| t.is_dir()).unwrap_or(false) {
                        if let Some(n) = e.file_name().to_str() {
                            langs.push(n.to_string());
                        }
                    }
                }
                if !langs.is_empty() {
                    break;
                }
            }
        }
        if langs.is_empty() {
            langs = vec![
                "rust".into(),
                "javascript".into(),
                "typescript".into(),
                "python".into(),
                "vb".into(),
                "csharp".into(),
                "text".into(),
            ];
        }
        langs.sort();
        let root_dir = std::env::current_dir().unwrap_or_else(|_| std::path::PathBuf::from("/"));
        let _root_uri = format!("file://{}", root_dir.to_string_lossy());
        Self {
            font_system: FontSystem::new(),
            swash_cache: SwashCache::new(),
            win_width: 0.0,
            win_height: 0.0,
            display_scale: SCALE,
            cursor: winit::window::CursorIcon::Default,
            tabs: Vec::new(),
            active_tab: 0,
            tree_view: TreeView::new(".", 2.0),
            all_languages: langs,
            current_lang: "rust".to_string(),
            lang_dropdown: None,
            theme_dropdown: None,
            clipboard: Clipboard::new().ok(),
            cmd_held: false,
            shift_held: false,
            alt_held: false,
            last_click_time: Instant::now(),
            click_count: 0,
            mouse_pos: (0.0, 0.0),
            explorer_width: EXPLORER_WIDTH,
            is_dragging_splitter: false,
            hovering_splitter: false,
            hovering_tab_close: None,
            is_dragging: false,
            needs_redraw: true,
            last_lsp_update: Instant::now(),
            pending_lsp_update: false,
            lsp: Arc::new(LspClient::new()),
            is_quick_open: false,
            quick_open_query: String::new(),
            is_command_palette: false,
            command_palette_query: String::new(),
            command_palette_selected: 0,
            is_project_search: false,
            project_search_query: String::new(),
            project_search_results: Vec::new(),
            project_search_selected: 0,
            tab_context_menu: None,
            tab_drag_idx: None,
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
                        Keybinding {
                            key: "z".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "Undo".into() },
                        Keybinding {
                            key: "z".into(),
                            cmd: true,
                            shift: true,
                            alt: false,
                            action: "Redo".into() },
                        Keybinding {
                            key: "a".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "SelectAll".into() },
                        Keybinding {
                            key: "s".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "Save".into() },
                        Keybinding {
                            key: "f".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "Find".into() },
                        Keybinding {
                            key: "h".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "Replace".into() },
                        Keybinding {
                            key: "/".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "ToggleComment".into() },
                        Keybinding {
                            key: "Tab".into(),
                            cmd: false,
                            shift: false,
                            alt: false,
                            action: "Indent".into() },
                        Keybinding {
                            key: "Tab".into(),
                            cmd: false,
                            shift: true,
                            alt: false,
                            action: "Unindent".into() },
                        Keybinding {
                            key: "`".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "ToggleOutput".into() },
                        Keybinding {
                            key: "ArrowUp".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "MoveBufferStart".into() },
                        Keybinding {
                            key: "ArrowDown".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "MoveBufferEnd".into() },
                        Keybinding {
                            key: "ArrowLeft".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "MoveLineStart".into() },
                        Keybinding {
                            key: "ArrowRight".into(),
                            cmd: true,
                            shift: false,
                            alt: false,
                            action: "MoveLineEnd".into() },
                        Keybinding {
                            key: "ArrowLeft".into(),
                            cmd: false,
                            shift: false,
                            alt: true,
                            action: "MoveWordLeft".into() },
                        Keybinding {
                            key: "ArrowRight".into(),
                            cmd: false,
                            shift: false,
                            alt: true,
                            action: "MoveWordRight".into() },
                    ];
                }
                kb
            },
            sidebar_tab: SidebarTab::Project,
            project: {
                let mut p = vybe_platform_dotnet::winforms::designer::project::Project::new("Project1".to_string());
                let mut form = vybe_platform_dotnet::winforms::designer::Form::new("Form1".to_string());
                form.width = 640;
                form.height = 480;
                p.forms
                    .push(vybe_platform_dotnet::winforms::designer::project::FormModule::new_classic(
                        form,
                    ));
                p.startup_object =
                    vybe_platform_dotnet::winforms::designer::project::StartupObject::Form("Form1".to_string());
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
            output_panel: OutputPanel::new(),
            output_panel_height: 150.0,
            output_lines_buffer: Vec::new(),
            goto_line_open: false,
            goto_line_query: String::new(),
            build_config: BuildConfig::Debug,
            split_panel: {
                let mut sp = SplitPanel::new(true); // horizontal split
                sp.set_split_pos(EXPLORER_WIDTH);
                sp
            },
            tab_panel: {
                let mut tp = TabPanel::new();
                tp.set_tab_height(TAB_BAR_HEIGHT);
                tp
            },
            status_bar: {
                let mut sb = StatusBarPanel::new();
                sb.set_height(FOOTER_HEIGHT);
                sb
            },
            sidebar_tabs: {
                let mut st = TabPanel::new();
                st.set_tab_height(SIDEBAR_TAB_H);
                st.set_tab_width(EXPLORER_WIDTH / 2.0);
                st.add_tab_header("Files", false);
                st.add_tab_header("Project", false);
                st.set_active(1); // default to Project tab
                st
            } }
    }

    /// Height in logical pixels of menu bar + toolbar (always present if any Form tab exists).
    pub(crate) fn top_chrome_h(&self) -> f32 {
        if self
            .tabs
            .iter()
            .any(|t| matches!(&t.content, TabContent::Form(_)))
        {
            28.0 + 36.0
        } else {
            0.0
        }
    }

    /// Sync toolkit containers with current window dimensions.
    pub(crate) fn relayout(&mut self) {
        let w = self.win_width;
        let h = self.win_height;
        let tch = self.top_chrome_h();

        // Status bar at bottom
        self.status_bar
            .set_rect(LayoutRect::new(0.0, h - FOOTER_HEIGHT, w, FOOTER_HEIGHT));

        // Split panel fills area between top chrome and status bar
        let main_h = h - tch - FOOTER_HEIGHT;
        self.split_panel
            .set_rect(LayoutRect::new(0.0, tch, w, main_h));
        self.split_panel.set_split_pos(self.explorer_width);

        // Tab panel occupies the right side of the split (panel2 area)
        let tab_x = self.explorer_width + SPLITTER_WIDTH + 1.0;
        let tab_w = (w - tab_x).max(0.0);
        let output_h = if self.output_panel.visible() {
            self.output_panel_height
        } else {
            0.0
        };
        self.tab_panel
            .set_rect(LayoutRect::new(tab_x, tch, tab_w, TAB_BAR_HEIGHT));

        // Output panel (above footer, right of sidebar)
        self.output_panel.set_rect(LayoutRect::new(
            tab_x,
            h - FOOTER_HEIGHT - output_h,
            tab_w,
            output_h,
        ));

        // Sidebar tabs at top of sidebar
        self.sidebar_tabs.set_rect(LayoutRect::new(
            0.0,
            tch,
            self.explorer_width,
            SIDEBAR_TAB_H,
        ));
        self.sidebar_tabs.set_tab_width(self.explorer_width / 2.0);
    }

    /// Sync tab_panel headers with the IDE tab list.
    pub(crate) fn sync_tab_headers(&mut self) {
        // Remove extra tabs from panel
        while self.tab_panel.tab_count() > self.tabs.len() {
            self.tab_panel.remove_tab(self.tab_panel.tab_count() - 1);
        }
        // Add missing tabs
        while self.tab_panel.tab_count() < self.tabs.len() {
            let i = self.tab_panel.tab_count();
            let closable = !self.tabs[i].is_sticky;
            self.tab_panel.add_tab_header(&self.tabs[i].name, closable);
        }
        // Update names and active index
        for (i, tab) in self.tabs.iter().enumerate() {
            let display_name = if tab.is_modified {
                format!("{}  \u{2022}", tab.name)
            } else {
                tab.name.clone()
            };
            self.tab_panel.set_tab_name(i, &display_name);
        }
        self.tab_panel.set_active(self.active_tab);
    }

    /// Process WidgetEvents from containers.
    pub(crate) fn process_widget_events(&mut self) {
        let tab_events = self.tab_panel.drain_events();
        for event in tab_events {
            match event {
                WidgetEvent::TabChanged(idx) => {
                    if idx < self.tabs.len() {
                        self.active_tab = idx;
                        let ext = self.tabs[idx]
                            .name
                            .rsplitn(2, '.')
                            .next()
                            .unwrap_or("")
                            .to_lowercase();
                        self.current_lang = App::lang_from_ext(&ext).to_string();
                    }
                }
                WidgetEvent::TabCloseRequested(idx) => {
                    if idx < self.tabs.len() && !self.tabs[idx].is_sticky {
                        self.tabs.remove(idx);
                        if self.active_tab >= self.tabs.len() && self.active_tab > 0 {
                            self.active_tab -= 1;
                        }
                        self.sync_tab_headers();
                    }
                }
                _ => {}
            }
        }
        let status_events = self.status_bar.drain_events();
        for event in status_events {
            match event {
                WidgetEvent::StatusBarClick(id) => match id.as_str() {
                    "lang" => {
                        let active_idx = self
                            .all_languages
                            .iter()
                            .position(|l| l == &self.current_lang)
                            .unwrap_or(0);
                        self.lang_dropdown = Some(Dropdown::new(
                            self.all_languages.clone(),
                            active_idx,
                            self.display_scale,
                            None,
                        ));
                    }
                    "theme" => {
                        let theme_names = vec![
                            "Silicon Green".into(),
                            "Cloud Blue".into(),
                            "Coffee Cream".into(),
                            "Sakura Pink".into(),
                            "One Dark".into(),
                            "Monokai".into(),
                            "GitHub Light".into(),
                            "Solarized Light".into(),
                            "Midnight".into(),
                            "Aura".into(),
                            "Veridian".into(),
                            "Rose".into(),
                            "Cyber".into(),
                            "Titanium".into(),
                            "Indigo Night".into(),
                        ];
                        let mut d = Dropdown::new(
                            theme_names,
                            self.current_theme_idx,
                            self.display_scale,
                            None,
                        );
                        d.num_cols = 2;
                        d.col_w = 160.0;
                        self.theme_dropdown = Some(d);
                    }
                    "config" => {
                        self.build_config = match self.build_config {
                            BuildConfig::Debug => BuildConfig::Release,
                            BuildConfig::Release => BuildConfig::Debug };
                    }
                    "diagnostics" => {
                        self.output_panel.set_visible(true);
                        self.output_panel
                            .set_active_tab(vybe_widgets::output_panel::OutputTab::Problems);
                    }
                    _ => {}
                },
                _ => {}
            }
        }
        // Sidebar tabs events
        let sidebar_events = self.sidebar_tabs.drain_events();
        for event in sidebar_events {
            match event {
                WidgetEvent::TabChanged(idx) => {
                    self.sidebar_tab = match idx {
                        0 => SidebarTab::Files,
                        _ => SidebarTab::Project };
                }
                _ => {}
            }
        }
        // Output panel events
        let output_events = self.output_panel.drain_panel_events();
        for event in output_events {
            match event {
                OutputPanelEvent::Close => {
                    self.output_panel.set_visible(false);
                }
                OutputPanelEvent::ClearOutput => {
                    self.output_panel.clear_output();
                    self.output_lines_buffer.clear();
                }
                OutputPanelEvent::TabChanged(tab) => {
                    self.output_panel.set_active_tab(tab);
                }
                OutputPanelEvent::ProblemClicked(idx) => {
                    // Build flat list of (tab_index, line) for all diagnostics
                    let mut diag_entries: Vec<(usize, usize)> = Vec::new();
                    for (ti, t) in self.tabs.iter().enumerate() {
                        if let TabContent::Code(cw) = &t.content {
                            for d in &cw.my_editor.diagnostics {
                                diag_entries.push((ti, d.line));
                            }
                        }
                    }
                    if let Some(&(tab_idx, diag_line)) = diag_entries.get(idx) {
                        self.active_tab = tab_idx;
                        if let TabContent::Code(cw) = &mut self.tabs[tab_idx].content {
                            let max_line = cw.line_count().saturating_sub(1);
                            let safe_line = diag_line.min(max_line);
                            cw.set_cursor_pos(safe_line, 0);
                            cw.needs_reshape = true;
                            cw.scroll_y = (safe_line as f32 * 20.0).max(0.0);
                        }
                    }
                }
            }
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
            _ => Theme::one_dark() }
    }

    pub fn get_theme_name(&self) -> &str {
        match self.current_theme_idx {
            0 => "Silicon Green",
            1 => "Cloud Blue",
            2 => "Coffee Cream",
            3 => "Sakura Pink",
            4 => "One Dark",
            5 => "Monokai",
            6 => "Frost Light",
            7 => "Solarized Light",
            8 => "Midnight",
            9 => "Aura",
            10 => "Veridian",
            11 => "Rose",
            12 => "Cyber",
            13 => "Titanium",
            14 => "Indigo Night",
            _ => "One Dark" }
    }

    fn draw_ui_text(
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        text: &str,
        x: f32,
        y: f32,
        col: TextColor,
    ) {
        crate::ide_text::draw_mono(pix, fs, sc, text, x / SCALE, y / SCALE, 14.0, col, SCALE);
    }

    /// Map a file extension to a language name understood by `load_language`.
    pub(crate) fn lang_from_ext(ext: &str) -> &'static str {
        match ext {
            "rs" => "rust",
            "js" | "mjs" | "cjs" => "javascript",
            "ts" | "tsx" => "typescript",
            "vb" | "bas" | "frm" => "vb",
            "cs" => "csharp",
            "py" => "python",
            "php" => "php",
            "rb" => "ruby",
            "dart" => "dart",
            "java" => "java",
            "go" => "go",
            "c" | "h" => "c",
            "cpp" | "cc" | "cxx" | "hpp" => "cpp",
            "html" | "htm" => "html",
            "css" => "css",
            "json" => "json",
            "yaml" | "yml" => "yaml",
            "md" => "markdown",
            "sql" => "sql",
            "sh" | "bash" => "shell",
            "ps" | "ps1" | "psm1" | "psd1" => "powershell",
            "lua" => "lua",
            "r" => "r",
            "f90" | "f95" | "f03" | "f08" | "for" | "f" => "fortran",
            "cob" | "cbl" => "cobol",
            _ => "plaintext" }
    }

    /// Build a file URI for a tab: uses `tab.path` when set, otherwise
    /// derives one from the current working directory.
    pub(crate) fn tab_uri(tab: &Tab) -> String {
        tab.path
            .as_deref()
            .map(|p| format!("file://{}", p))
            .unwrap_or_else(|| {
                let cwd = std::env::current_dir()
                    .map(|p| p.to_string_lossy().into_owned())
                    .unwrap_or_default();
                format!("file://{}/{}", cwd, tab.name)
            })
    }
}

// ── Application (toolkit) ──────────────────────────────────────────────

impl vybe_widgets::Application for App {
    fn title(&self) -> String {
        let tab_part = self
            .tabs
            .get(self.active_tab)
            .map(|t| {
                let dot = if t.is_modified { " •" } else { "" };
                format!("{}{} — ", t.name, dot)
            })
            .unwrap_or_default();
        let project = &self.project.name;
        format!("{}{}  —  Vybe IDE", tab_part, project)
    }

    fn on_init(&mut self, width: f32, height: f32, scale: f32) {
        self.win_width = width;
        self.win_height = height;
        self.display_scale = scale;

        // Open with the Form Designer tab only — no scratch file
        let mut designer_state = crate::form_designer_tab::FormDesignerState::new();
        if let Some(fm) = self.project.forms.first() {
            designer_state.form = fm.form.clone();
        }
        self.tabs.push(Tab {
            name: "Form Designer".to_string(),
            path: None,
            content: TabContent::Form(designer_state),
            is_sticky: true,
            buffer: None,
            is_modified: false });
        self.active_tab = 0;
        self.sync_tab_headers();
        self.relayout();
    }

    fn on_resize(&mut self, width: f32, height: f32) {
        self.win_width = width;
        self.win_height = height;
        self.relayout();
        self.sync_tab_headers();
    }

    fn render(&mut self, pix: &mut Pixmap, _scale: f32) {
        self.render_internal(pix);
    }

    fn handle_mouse(&mut self, event: vybe_widgets::MouseEvent) -> bool {
        use vybe_widgets::layout::{MouseButton as WMouseButton, MouseEventKind};
        self.cmd_held = event.cmd;
        self.shift_held = event.shift;
        self.alt_held = event.alt;
        self.mouse_pos = (event.x * SCALE, event.y * SCALE);

        // Route tab bar events through the TabPanel
        if self.tab_panel.rect().contains(event.x, event.y) {
            if self.tab_panel.handle_mouse(&event) {
                self.process_widget_events();
                return true;
            }
        }

        // Route status bar events through StatusBarPanel
        if self.status_bar.rect().contains(event.x, event.y) {
            if self.status_bar.handle_mouse(&event) {
                self.process_widget_events();
                return true;
            }
        }

        // Fall through to legacy handlers
        match event.kind {
            MouseEventKind::Move => {
                self.handle_cursor_moved();
            }
            MouseEventKind::Press(btn) | MouseEventKind::Release(btn) => {
                let winit_btn = match btn {
                    WMouseButton::Left => winit::event::MouseButton::Left,
                    WMouseButton::Right => winit::event::MouseButton::Right,
                    WMouseButton::Middle => winit::event::MouseButton::Middle };
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

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        // Tab bar scrolling via toolkit
        let tab_rect = self.tab_panel.rect();
        if tab_rect.contains(x, y) {
            self.tab_panel.scroll_tab_bar(delta);
            return true;
        }

        // Fall through to legacy scroll handler
        let winit_delta = winit::event::MouseScrollDelta::PixelDelta(
            winit::dpi::PhysicalPosition::new(0.0, delta as f64 / 2.0),
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
        self.is_quick_open = false;
        self.is_command_palette = false;
        self.is_project_search = false;
        self.goto_line_open = false;
        self.tab_context_menu = None;
        self.pe_context_menu = None;
        self.lang_dropdown = None;
        self.theme_dropdown = None;
    }
}

// ── Entry Point ────────────────────────────────────────────────────────

pub fn run_gui(my_editor: MyEditor, open_form: bool) {
    let app = App::new(my_editor, open_form);
    vybe_widgets::run_app("Vybe IDE", 1200, 900, SCALE, app);
}
