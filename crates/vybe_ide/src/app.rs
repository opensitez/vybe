//! Main IDE application state and rendering logic.

use cosmic_text::{FontSystem, SwashCache};
use tiny_skia::Pixmap;
use vybe_forms::Form;
use vybe_project::project::{Project, FormModule, CodeFile, StartupObject};

use crate::layout::{IdeLayout, LayoutConfig};
use crate::panels::menu_bar::{MenuBar, MenuAction};
use crate::panels::toolbar::{Toolbar, ToolbarAction};
use crate::panels::project_explorer::{ProjectExplorer, ExplorerEvent};
use crate::panels::toolbox_panel::{ToolboxPanel, ControlTool};
use crate::panels::properties_panel::PropertiesPanel;
use crate::panels::form_designer::FormDesigner;
use crate::panels::code_editor::CodeEditor;
use crate::panels::status_bar::StatusBar;
use crate::panels::project_properties::ProjectPropertiesDialog;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CenterView {
    FormDesigner,
    CodeEditor,
}

pub struct SkiaIde {
    pub fs: FontSystem,
    pub sc: SwashCache,

    pub project: Project,
    pub current_form: Option<String>,
    pub center_view: CenterView,
    pub project_path: Option<String>,
    pub clipboard: Vec<vybe_forms::Control>,

    pub layout_config: LayoutConfig,
    pub scale: f32,
    pub win_w: f32,
    pub win_h: f32,

    pub menu_bar: MenuBar,
    pub toolbox: ToolboxPanel,
    pub explorer: ProjectExplorer,
    pub properties: PropertiesPanel,
    pub form_designer: FormDesigner,
    pub code_editor: CodeEditor,
    pub status_bar: StatusBar,

    pub mouse_down: bool,
    pub dragging_splitter: bool,
    pub project_props_dialog: ProjectPropertiesDialog,
}

impl SkiaIde {
    pub fn new(scale: f32) -> Self {
        let mut project = Project::new("Project1");
        let mut form = Form::new("Form1");
        form.width = 640;
        form.height = 480;
        let fm = FormModule::new_classic(form);
        project.forms.push(fm);
        project.startup_object = StartupObject::Form("Form1".to_string());

        Self {
            fs: FontSystem::new(),
            sc: SwashCache::new(),
            project,
            current_form: Some("Form1".to_string()),
            center_view: CenterView::FormDesigner,
            project_path: None,
            clipboard: Vec::new(),
            layout_config: LayoutConfig::default(),
            scale,
            win_w: 1200.0,
            win_h: 800.0,
            menu_bar: MenuBar::new(),
            toolbox: ToolboxPanel::new(),
            explorer: ProjectExplorer::new(),
            properties: PropertiesPanel::new(),
            form_designer: FormDesigner::new(),
            code_editor: CodeEditor::new(),
            status_bar: StatusBar::new(),
            mouse_down: false,
            dragging_splitter: false,
            project_props_dialog: ProjectPropertiesDialog::new(),
        }
    }

    /// Find the current form index by name.
    fn current_form_idx(&self) -> Option<usize> {
        let name = self.current_form.as_ref()?;
        self.project.forms.iter().position(|f| &f.form.name == name)
    }

    /// Get the current form.
    fn current_form_ref(&self) -> Option<&Form> {
        self.current_form_idx().and_then(|i| self.project.forms.get(i)).map(|fm| &fm.form)
    }

    /// Whether we're in form designer mode (not code, and a form is selected).
    fn in_form_designer(&self) -> bool {
        self.center_view == CenterView::FormDesigner && self.current_form_ref().is_some()
    }

    pub fn render(&mut self, pix: &mut Pixmap) {
        let layout = IdeLayout::compute(self.win_w, self.win_h, &self.layout_config);
        let scale = self.scale;
        let in_designer = self.in_form_designer();

        pix.fill(tiny_skia::Color::from_rgba8(240, 240, 240, 255));

        self.menu_bar.render(pix, &mut self.fs, &mut self.sc, layout.menu_bar, scale);
        Toolbar::render(pix, &mut self.fs, &mut self.sc, layout.toolbar, scale);

        if self.layout_config.show_project_explorer && layout.project_explorer.w > 0.0 {
            self.explorer.render(
                pix, &mut self.fs, &mut self.sc,
                layout.project_explorer, scale,
                &self.project, self.current_form.as_deref(),
            );
        }

        // Splitter between explorer and toolbox
        if layout.splitter.h > 0.0 && in_designer {
            let sr = layout.splitter;
            let mut sp = tiny_skia::Paint::default();
            sp.set_color_rgba8(210, 210, 210, 255);
            if let Some(r) = tiny_skia::Rect::from_xywh(sr.x * scale, sr.y * scale, sr.w * scale, sr.h * scale) {
                pix.fill_rect(r, &sp, tiny_skia::Transform::identity(), None);
            }
            // Grip dots
            sp.set_color_rgba8(160, 160, 160, 255);
            let cy = sr.y + sr.h / 2.0;
            for i in 0..3 {
                let dx = sr.w / 2.0 - 8.0 + i as f32 * 8.0;
                if let Some(r) = tiny_skia::Rect::from_xywh((sr.x + dx) * scale, (cy - 0.5) * scale, 2.0 * scale, 2.0 * scale) {
                    pix.fill_rect(r, &sp, tiny_skia::Transform::identity(), None);
                }
            }
        }

        // Toolbox only in form designer mode
        if in_designer && self.layout_config.show_toolbox && layout.toolbox.w > 0.0 {
            self.toolbox.render(pix, &mut self.fs, &mut self.sc, layout.toolbox, scale);
        }

        match self.center_view {
            CenterView::FormDesigner => {
                if let Some(idx) = self.current_form_idx() {
                    if let Some(fm) = self.project.forms.get(idx) {
                        self.form_designer.render(
                            pix, &mut self.fs, &mut self.sc,
                            layout.center, scale, &fm.form,
                        );
                    }
                }
            }
            CenterView::CodeEditor => {
                self.code_editor.render(pix, &mut self.fs, &mut self.sc, layout.center, scale);
            }
        }

        // Properties only in form designer mode
        if in_designer && self.layout_config.show_properties && layout.properties.w > 0.0 {
            let form = self.current_form_idx()
                .and_then(|i| self.project.forms.get(i))
                .map(|fm| &fm.form);
            let selected = form.and_then(|f| self.form_designer.selected_control_name(f));
            self.properties.render(
                pix, &mut self.fs, &mut self.sc,
                layout.properties, scale, form, selected,
            );
        }

        self.status_bar.render(pix, &mut self.fs, &mut self.sc, layout.status_bar, scale);

        // Menu dropdown overlay — rendered LAST so it draws on top of everything
        self.menu_bar.render_dropdown_overlay(pix, &mut self.fs, &mut self.sc, layout.menu_bar, scale);

        // Project properties dialog (modal overlay)
        self.project_props_dialog.render(
            pix, &mut self.fs, &mut self.sc,
            self.win_w, self.win_h, scale, &self.project,
        );
    }

    pub fn handle_mouse_down(&mut self, lx: f32, ly: f32, ctrl_held: bool) {
        self.mouse_down = true;

        // Project properties dialog (modal — blocks everything else)
        if self.project_props_dialog.visible {
            if self.project_props_dialog.is_ok_clicked(lx, ly, self.win_w, self.win_h) {
                self.project_props_dialog.apply(&mut self.project);
                self.project_props_dialog.close();
                self.status_bar.message = "Project properties updated".to_string();
            } else {
                let project_clone = self.project.clone();
                self.project_props_dialog.handle_click(lx, ly, self.win_w, self.win_h, &project_clone);
            }
            return;
        }

        let layout = IdeLayout::compute(self.win_w, self.win_h, &self.layout_config);

        // Menu bar (or open dropdown)
        if layout.menu_bar.contains(lx, ly) || self.menu_bar.open_menu.is_some() {
            if let Some(action) = self.menu_bar.handle_click(lx, ly, layout.menu_bar) {
                self.handle_action(action);
            }
            return;
        }
        self.menu_bar.open_menu = None;

        // Toolbar
        if layout.toolbar.contains(lx, ly) {
            if let Some(action) = Toolbar::handle_click(lx, ly, layout.toolbar) {
                self.handle_toolbar_action(action);
            }
            return;
        }

        // Splitter drag
        if layout.splitter.h > 0.0 && layout.splitter.contains(lx, ly) {
            self.dragging_splitter = true;
            return;
        }

        // Project explorer
        if layout.project_explorer.contains(lx, ly) && self.layout_config.show_project_explorer {
            let project_clone = self.project.clone();
            match self.explorer.handle_click(lx, ly, layout.project_explorer, &project_clone) {
                ExplorerEvent::SelectForm(name) => {
                    self.current_form = Some(name.clone());
                    self.center_view = CenterView::FormDesigner;
                    self.form_designer.selected_controls.clear();
                    self.status_bar.message = format!("Form: {}", name);
                }
                ExplorerEvent::SelectCode(name) => {
                    // Load code file content into editor
                    if let Some(cf) = self.project.code_files.iter().find(|c| c.name == name) {
                        self.code_editor.set_code(&cf.code);
                    }
                    self.current_form = Some(name.clone());
                    self.center_view = CenterView::CodeEditor;
                    self.status_bar.message = format!("Code: {}", name);
                }
                ExplorerEvent::ViewCode(name) => {
                    if let Some(fm) = self.project.forms.iter().find(|f| f.form.name == name) {
                        self.code_editor.set_code(fm.get_user_code());
                    }
                    self.center_view = CenterView::CodeEditor;
                    self.status_bar.message = format!("Code: {}", name);
                }
                ExplorerEvent::None => {}
            }
            return;
        }

        // Toolbox (only in designer mode)
        if self.in_form_designer() && layout.toolbox.contains(lx, ly) && self.layout_config.show_toolbox {
            self.toolbox.handle_click(lx, ly, layout.toolbox);
            return;
        }

        // Properties panel (only in designer mode)
        if self.in_form_designer() && layout.properties.contains(lx, ly) && self.layout_config.show_properties {
            let form = self.current_form_idx().and_then(|i| self.project.forms.get(i)).map(|fm| &fm.form);
            let selected = form.and_then(|f| self.form_designer.selected_control_name(f));
            self.properties.handle_click(lx, ly, layout.properties, form, selected);

            // ── Handle immediate commits for inline toggles / pickers ──
            if self.properties.pending_commit {
                self.commit_property_edit();
                self.properties.pending_commit = false;
            }

            // ── Handle pending event (click on Events tab row) ──
            if let Some((ctrl_name, event_name)) = self.properties.pending_event.take() {
                self.handle_event_click(&ctrl_name, &event_name);
            }

            // ── Handle pending connection wizard action ──
            if let Some(action) = self.properties.pending_action.take() {
                self.handle_conn_action(&action);
            }

            return;
        }

        // Commit any active property edit
        if self.properties.editing.is_some() {
            self.commit_property_edit();
        }

        // Center area
        match self.center_view {
            CenterView::FormDesigner => {
                if let Some(idx) = self.current_form_idx() {
                    let tool = self.toolbox.selected_tool();
                    if let ControlTool::Control(ref ct) = tool {
                        if ct.is_non_visual() {
                            // Non-visual controls go straight to the tray
                            if let Some(fm) = self.project.forms.get_mut(idx) {
                                self.form_designer.add_non_visual(ct.clone(), &mut fm.form);
                                self.toolbox.reset_to_pointer();
                                self.status_bar.message = "Component added".to_string();
                                return;
                            }
                        } else if let Some(fm) = self.project.forms.get_mut(idx) {
                            if self.form_designer.place_control(lx, ly, layout.center, &mut fm.form, tool) {
                                self.toolbox.reset_to_pointer();
                                self.status_bar.message = "Control placed".to_string();
                                return;
                            }
                        }
                    }
                    if let Some(fm) = self.project.forms.get(idx) {
                        let form_ref = fm.form.clone();
                        self.form_designer.handle_mouse_down(lx, ly, layout.center, &form_ref, ctrl_held);
                    }
                }
            }
            CenterView::CodeEditor => {
                self.code_editor.handle_click(lx, ly, layout.center);
            }
        }
    }

    pub fn handle_mouse_hover(&mut self, lx: f32, ly: f32) {
        let layout = IdeLayout::compute(self.win_w, self.win_h, &self.layout_config);
        self.menu_bar.handle_hover(lx, ly, layout.menu_bar);
    }

    pub fn handle_mouse_move(&mut self, lx: f32, ly: f32) {
        if !self.mouse_down { return; }

        // Splitter drag
        if self.dragging_splitter {
            let body_top = 28.0 + 36.0; // MENU_H + TOOLBAR_H
            let body_h = (self.win_h - body_top - 24.0).max(1.0);
            let frac = ((ly - body_top) / body_h).clamp(0.15, 0.85);
            self.layout_config.left_split = frac;
            return;
        }

        let layout = IdeLayout::compute(self.win_w, self.win_h, &self.layout_config);
        if self.center_view == CenterView::FormDesigner {
            if let Some(idx) = self.current_form_idx() {
                if let Some(fm) = self.project.forms.get_mut(idx) {
                    self.form_designer.handle_mouse_move(lx, ly, layout.center, &mut fm.form);
                }
            }
        }
    }

    pub fn handle_mouse_up(&mut self, _lx: f32, _ly: f32) {
        self.mouse_down = false;
        self.dragging_splitter = false;
        let layout = IdeLayout::compute(self.win_w, self.win_h, &self.layout_config);
        if self.center_view == CenterView::FormDesigner {
            if let Some(idx) = self.current_form_idx() {
                if let Some(fm) = self.project.forms.get(idx) {
                    self.form_designer.handle_mouse_up(layout.center, &fm.form);
                }
            }
        }
    }

    pub fn handle_scroll(&mut self, delta: f32, lx: f32, ly: f32) {
        let layout = IdeLayout::compute(self.win_w, self.win_h, &self.layout_config);
        if layout.project_explorer.contains(lx, ly) && self.layout_config.show_project_explorer {
            self.explorer.scroll(delta, layout.project_explorer, &self.project);
            return;
        }
        if layout.toolbox.contains(lx, ly) && self.layout_config.show_toolbox && self.in_form_designer() {
            self.toolbox.scroll(delta, layout.toolbox);
            return;
        }
        if layout.properties.contains(lx, ly) && self.layout_config.show_properties {
            let form = self.current_form_idx().and_then(|i| self.project.forms.get(i)).map(|fm| &fm.form);
            let selected = form.and_then(|f| self.form_designer.selected_control_name(f));
            self.properties.scroll(delta, layout.properties, form, selected);
            return;
        }
        if layout.center.contains(lx, ly) {
            match self.center_view {
                CenterView::CodeEditor => self.code_editor.scroll(delta, layout.center),
                CenterView::FormDesigner => {
                    self.form_designer.scroll_y = (self.form_designer.scroll_y - delta * 20.0).max(0.0);
                }
            }
        }
    }

    pub fn handle_key(&mut self, key: &str) {
        if self.properties.editing.is_some() {
            if self.properties.handle_key(key) {
                self.commit_property_edit();
            }
            return;
        }
        // Escape — close dialog, menu, or deselect
        if key == "Escape" {
            if self.project_props_dialog.visible {
                self.project_props_dialog.close();
                return;
            }
            if self.menu_bar.open_menu.is_some() {
                self.menu_bar.open_menu = None;
            } else if !self.form_designer.selected_controls.is_empty() {
                self.form_designer.selected_controls.clear();
            }
            return;
        }
        // Delete selected controls
        if (key == "Delete" || key == "Backspace") && self.center_view == CenterView::FormDesigner {
            if !self.form_designer.selected_controls.is_empty() {
                if let Some(idx) = self.current_form_idx() {
                    if let Some(fm) = self.project.forms.get_mut(idx) {
                        let sel = &self.form_designer.selected_controls;
                        fm.form.controls.retain(|c| !sel.contains(&c.id));
                        self.form_designer.selected_controls.clear();
                        self.status_bar.message = "Deleted".to_string();
                    }
                }
                return;
            }
        }
        if self.center_view == CenterView::CodeEditor {
            self.code_editor.handle_key(key);
        }
    }

    pub fn handle_char(&mut self, ch: char) {
        if self.properties.editing.is_some() {
            self.properties.handle_char(ch);
            return;
        }
        if self.center_view == CenterView::CodeEditor && !ch.is_control() {
            self.code_editor.insert_char(ch);
        }
    }

    pub fn handle_shortcut(&mut self, shortcut: &str) {
        match shortcut {
            "copy" => {
                if self.center_view == CenterView::FormDesigner {
                    if let Some(idx) = self.current_form_idx() {
                        if let Some(fm) = self.project.forms.get(idx) {
                            let sel = &self.form_designer.selected_controls;
                            self.clipboard = fm.form.controls.iter()
                                .filter(|c| sel.contains(&c.id))
                                .cloned()
                                .collect();
                            self.status_bar.message = format!("Copied {} control(s)", self.clipboard.len());
                        }
                    }
                }
            }
            "cut" => {
                self.handle_shortcut("copy");
                // Delete selected
                if let Some(idx) = self.current_form_idx() {
                    if let Some(fm) = self.project.forms.get_mut(idx) {
                        let sel = &self.form_designer.selected_controls;
                        fm.form.controls.retain(|c| !sel.contains(&c.id));
                        self.form_designer.selected_controls.clear();
                        self.status_bar.message = "Cut".to_string();
                    }
                }
            }
            "paste" => {
                if self.center_view == CenterView::FormDesigner && !self.clipboard.is_empty() {
                    if let Some(idx) = self.current_form_idx() {
                        if let Some(fm) = self.project.forms.get_mut(idx) {
                            let mut new_ids = Vec::new();
                            for orig in &self.clipboard {
                                let mut ctrl = orig.clone();
                                ctrl.id = uuid::Uuid::new_v4();
                                ctrl.bounds.x += 20;
                                ctrl.bounds.y += 20;
                                // Rename to avoid duplicates
                                let base = ctrl.name.trim_end_matches(|c: char| c.is_ascii_digit()).to_string();
                                let mut n = 1u32;
                                loop {
                                    let candidate = format!("{}{}", base, n);
                                    if !fm.form.controls.iter().any(|c| c.name == candidate) {
                                        ctrl.name = candidate;
                                        break;
                                    }
                                    n += 1;
                                }
                                new_ids.push(ctrl.id);
                                fm.form.controls.push(ctrl);
                            }
                            self.form_designer.selected_controls = new_ids;
                            self.status_bar.message = "Pasted".to_string();
                        }
                    }
                }
            }
            "select_all" => {
                if self.center_view == CenterView::FormDesigner {
                    if let Some(idx) = self.current_form_idx() {
                        if let Some(fm) = self.project.forms.get(idx) {
                            self.form_designer.selected_controls = fm.form.controls.iter()
                                .filter(|c| !c.control_type.is_non_visual())
                                .map(|c| c.id)
                                .collect();
                        }
                    }
                }
            }
            "undo" => self.handle_action(MenuAction::Undo),
            "redo" => self.handle_action(MenuAction::Redo),
            "save" => self.handle_action(MenuAction::SaveProject),
            "new" => self.handle_action(MenuAction::NewProject),
            "open" => self.handle_action(MenuAction::OpenProject),
            _ => {}
        }
    }

    fn commit_property_edit(&mut self) {
        if let Some((key, value)) = self.properties.commit_edit() {
            let key_display = key.clone();
            if let Some(idx) = self.current_form_idx() {
                if let Some(fm) = self.project.forms.get_mut(idx) {
                    let selected_name = self.form_designer.selected_control_name(&fm.form)
                        .map(|s| s.to_string());
                    if let Some(ctrl_name) = selected_name {
                        if let Some(ctrl) = fm.form.controls.iter_mut().find(|c| c.name == ctrl_name) {
                            match key.as_str() {
                                "Name" => { ctrl.name = value; }
                                "Left" | "X" => { if let Ok(v) = value.parse() { ctrl.bounds.x = v; } }
                                "Top" | "Y" => { if let Ok(v) = value.parse() { ctrl.bounds.y = v; } }
                                "Width" => { if let Ok(v) = value.parse() { ctrl.bounds.width = v; } }
                                "Height" => { if let Ok(v) = value.parse() { ctrl.bounds.height = v; } }
                                "TabIndex" => { if let Ok(v) = value.parse() { ctrl.tab_index = v; } }
                                // Connection builder fields → auto-build ConnectionString
                                "DbType" | "DbPath" | "DbHost" | "DbPort" | "DbName" | "DbUser" | "DbPassword" => {
                                    ctrl.properties.set(key.clone(), value);
                                    // Rebuild connection string
                                    let db_type = ctrl.properties.get_string("DbType").unwrap_or("SQLite").to_string();
                                    let conn = match db_type.as_str() {
                                        "SQLite" => {
                                            let path = ctrl.properties.get_string("DbPath").unwrap_or("database.db");
                                            format!("Data Source={}", path)
                                        }
                                        "PostgreSQL" => {
                                            let host = ctrl.properties.get_string("DbHost").unwrap_or("localhost");
                                            let port = ctrl.properties.get_string("DbPort").unwrap_or("5432");
                                            let db = ctrl.properties.get_string("DbName").unwrap_or("");
                                            let user = ctrl.properties.get_string("DbUser").unwrap_or("postgres");
                                            let pass = ctrl.properties.get_string("DbPassword").unwrap_or("");
                                            format!("Host={};Port={};Database={};Username={};Password={}", host, port, db, user, pass)
                                        }
                                        "MySQL" => {
                                            let host = ctrl.properties.get_string("DbHost").unwrap_or("localhost");
                                            let port = ctrl.properties.get_string("DbPort").unwrap_or("3306");
                                            let db = ctrl.properties.get_string("DbName").unwrap_or("");
                                            let user = ctrl.properties.get_string("DbUser").unwrap_or("root");
                                            let pass = ctrl.properties.get_string("DbPassword").unwrap_or("");
                                            format!("Server={};Port={};Database={};Uid={};Pwd={}", host, port, db, user, pass)
                                        }
                                        _ => String::new(),
                                    };
                                    ctrl.properties.set("ConnectionString", conn);
                                }
                                _ => { ctrl.properties.set(key, value); }
                            }
                        }
                    } else {
                        match key.as_str() {
                            "Name" => { fm.form.name = value; }
                            "Text" => { fm.form.text = value; }
                            "Width" => { if let Ok(v) = value.parse() { fm.form.width = v; } }
                            "Height" => { if let Ok(v) = value.parse() { fm.form.height = v; } }
                            "BackColor" => { fm.form.back_color = if value.is_empty() { None } else { Some(value) }; }
                            _ => {}
                        }
                    }
                    self.status_bar.message = format!("{} updated", key_display);
                }
            }
        }
    }

    fn handle_toolbar_action(&mut self, action: ToolbarAction) {
        match action {
            ToolbarAction::Run => {
                self.status_bar.message = "Running...".to_string();
            }
            ToolbarAction::Stop => {
                self.status_bar.message = "Stopped".to_string();
            }
            ToolbarAction::ViewDesigner => {
                self.center_view = CenterView::FormDesigner;
                self.status_bar.message = "Designer view".to_string();
            }
            ToolbarAction::ViewCode => {
                if let Some(idx) = self.current_form_idx() {
                    if let Some(fm) = self.project.forms.get(idx) {
                        self.code_editor.set_code(fm.get_user_code());
                    }
                }
                self.center_view = CenterView::CodeEditor;
                self.status_bar.message = "Code view".to_string();
            }
            ToolbarAction::AddForm => self.handle_action(MenuAction::AddForm),
            ToolbarAction::AddCode => self.handle_action(MenuAction::AddModule),
        }
    }

    fn handle_action(&mut self, action: MenuAction) {
        match action {
            MenuAction::NewProject => {
                self.project = Project::new("Project1");
                let mut form = Form::new("Form1");
                form.width = 640;
                form.height = 480;
                let fm = FormModule::new_classic(form);
                self.project.forms.push(fm);
                self.project.startup_object = StartupObject::Form("Form1".to_string());
                self.current_form = Some("Form1".to_string());
                self.center_view = CenterView::FormDesigner;
                self.form_designer.selected_controls.clear();
                self.project_path = None;
                self.status_bar.message = "New project created".to_string();
            }
            MenuAction::OpenProject => {
                if let Some(path) = rfd::FileDialog::new()
                    .add_filter("VB Project", &["vbproj", "vbp"])
                    .pick_file()
                {
                    let path_str = path.to_string_lossy().to_string();
                    match vybe_project::serialization::load_project_auto(&path_str) {
                        Ok(proj) => {
                            let first = proj.forms.first().map(|f| f.form.name.clone());
                            self.project = proj;
                            self.current_form = first;
                            self.center_view = CenterView::FormDesigner;
                            self.project_path = Some(path_str.clone());
                            self.form_designer.selected_controls.clear();
                            self.status_bar.message = format!("Opened: {}", path_str);
                        }
                        Err(e) => {
                            self.status_bar.message = format!("Error: {}", e);
                        }
                    }
                }
            }
            MenuAction::SaveProject => {
                if let Some(ref path) = self.project_path {
                    match vybe_project::serialization::save_project_auto(&self.project,path) {
                        Ok(_) => self.status_bar.message = format!("Saved: {}", path),
                        Err(e) => self.status_bar.message = format!("Save error: {}", e),
                    }
                } else {
                    self.handle_action(MenuAction::SaveAs);
                }
            }
            MenuAction::SaveAs => {
                if let Some(path) = rfd::FileDialog::new()
                    .add_filter("VB Project", &["vbproj"])
                    .save_file()
                {
                    let path_str = path.to_string_lossy().to_string();
                    match vybe_project::serialization::save_project_auto(&self.project,&path_str) {
                        Ok(_) => {
                            self.project_path = Some(path_str.clone());
                            self.status_bar.message = format!("Saved: {}", path_str);
                        }
                        Err(e) => self.status_bar.message = format!("Save error: {}", e),
                    }
                }
            }
            MenuAction::Exit => {
                std::process::exit(0);
            }
            MenuAction::AddForm => {
                let name = format!("Form{}", self.project.forms.len() + 1);
                let mut form = Form::new(&name);
                form.width = 640;
                form.height = 480;
                let fm = FormModule::new_classic(form);
                self.project.forms.push(fm);
                self.current_form = Some(name.clone());
                self.center_view = CenterView::FormDesigner;
                self.form_designer.selected_controls.clear();
                self.status_bar.message = format!("Added form: {}", name);
            }
            MenuAction::AddModule => {
                let name = format!("Module{}.vb", self.project.code_files.len() + 1);
                self.project.code_files.push(CodeFile {
                    name: name.clone(),
                    code: format!("Module {}\n\nEnd Module\n", name.replace(".vb", "")),
                });
                self.code_editor.set_code(&self.project.code_files.last().unwrap().code);
                self.current_form = Some(name.clone());
                self.center_view = CenterView::CodeEditor;
                self.status_bar.message = format!("Added module: {}", name);
            }
            MenuAction::RunProject => {
                self.status_bar.message = "Run not yet implemented".to_string();
            }
            MenuAction::StopProject => {
                self.status_bar.message = "Stopped".to_string();
            }
            MenuAction::Undo => {
                self.status_bar.message = "Undo not yet implemented".to_string();
            }
            MenuAction::Redo => {
                self.status_bar.message = "Redo not yet implemented".to_string();
            }
            MenuAction::Cut | MenuAction::Copy | MenuAction::Paste => {
                self.status_bar.message = format!("{:?} not yet implemented", action);
            }
            MenuAction::About => {
                // "Project Properties..." in the legacy editor is mapped to About action
                self.project_props_dialog.open(&self.project);
            }
        }
    }

    /// Handle clicking on an event row: open existing handler or create skeleton.
    fn handle_event_click(&mut self, ctrl_name: &str, event_name: &str) {
        let form_name = self.current_form.clone().unwrap_or_default();
        let handler_name = format!("{}_{}", ctrl_name, event_name);

        // Determine event parameter signature
        let params = Self::event_params(event_name);

        // Get current code
        let code = if let Some(idx) = self.current_form_idx() {
            self.project.forms.get(idx).map(|fm| fm.get_user_code().to_string()).unwrap_or_default()
        } else {
            String::new()
        };

        // Check if handler already exists (simple string search)
        let sub_decl_handles = format!("Private Sub {}({}) Handles {}.{}",
            handler_name, params, ctrl_name, event_name);
        let sub_decl_plain = format!("Private Sub {}({})", handler_name, params);

        let handler_exists = code.contains(&sub_decl_handles) || code.contains(&sub_decl_plain)
            || code.to_lowercase().contains(&handler_name.to_lowercase());

        if handler_exists {
            // Handler exists — switch to code editor
            self.code_editor.set_code(&code);
            self.center_view = CenterView::CodeEditor;
            self.status_bar.message = format!("Opened handler: {}", handler_name);
        } else {
            // Generate skeleton handler
            let is_form_event = ctrl_name == form_name;
            let handles_clause = if is_form_event {
                format!("Handles Me.{}", event_name)
            } else {
                format!("Handles {}.{}", ctrl_name, event_name)
            };

            let snippet = format!(
                "    Private Sub {}({}) {}\n        ' TODO: Add your code here\n    End Sub",
                handler_name, params, handles_clause
            );

            // Insert before "End Class" if present, otherwise append
            let new_code = Self::insert_before_end_class(&code, &snippet);

            // Update the form module's code
            if let Some(idx) = self.current_form_idx() {
                if let Some(fm) = self.project.forms.get_mut(idx) {
                    fm.set_user_code(new_code.clone());
                }
            }

            self.code_editor.set_code(&new_code);
            self.center_view = CenterView::CodeEditor;
            self.status_bar.message = format!("Created handler: {}", handler_name);
        }
    }

    /// Insert a code snippet before "End Class", or append.
    fn insert_before_end_class(code: &str, snippet: &str) -> String {
        let lower = code.to_lowercase();
        if let Some(idx) = lower.rfind("end class") {
            let (head, tail) = code.split_at(idx);
            format!("{}\n\n{}\n{}", head.trim_end(), snippet, tail)
        } else {
            format!("{}\n\n{}", code, snippet)
        }
    }

    /// Get event parameter signature for common event types.
    fn event_params(event_name: &str) -> &'static str {
        match event_name.to_lowercase().as_str() {
            "mouseclick" | "mousedoubleclick" | "mousedown" | "mouseup" | "mousemove" | "mousewheel"
            | "nodemouseclick" | "nodemousedoubleclick" => "sender As Object, e As MouseEventArgs",
            "keydown" | "keyup" => "sender As Object, e As KeyEventArgs",
            "keypress" => "sender As Object, e As KeyPressEventArgs",
            "formclosing" => "sender As Object, e As FormClosingEventArgs",
            "formclosed" => "sender As Object, e As FormClosedEventArgs",
            "paint" | "cellpainting" => "sender As Object, e As PaintEventArgs",
            "cellclick" | "celldoubleclick" | "cellcontentclick" | "cellvaluechanged"
            | "cellendedit" | "cellbeginedit" | "cellvalidating" | "cellenter" | "cellleave"
            | "cellformatting" => "sender As Object, e As DataGridViewCellEventArgs",
            "afterselect" | "beforeselect" | "aftercheck" | "beforecheck"
            | "afterexpand" | "aftercollapse" | "beforeexpand" | "beforecollapse"
                => "sender As Object, e As TreeViewEventArgs",
            "scroll" => "sender As Object, e As ScrollEventArgs",
            "splittermoved" | "splittermoving" => "sender As Object, e As SplitterEventArgs",
            "linkclicked" => "sender As Object, e As LinkLabelLinkClickedEventArgs",
            "columnclick" => "sender As Object, e As ColumnClickEventArgs",
            _ => "sender As Object, e As EventArgs",
        }
    }

    /// Handle connection wizard actions: build connection string or test connection.
    fn handle_conn_action(&mut self, action: &str) {
        let selected_name = {
            let form = self.current_form_idx().and_then(|i| self.project.forms.get(i)).map(|fm| &fm.form);
            form.and_then(|f| self.form_designer.selected_control_name(f))
                .map(|s| s.to_string())
        };
        let ctrl_name = match selected_name {
            Some(n) => n,
            None => { self.status_bar.message = "No control selected".to_string(); return; }
        };

        match action {
            "build_conn" => {
                // Build connection string from individual DB fields
                if let Some(idx) = self.current_form_idx() {
                    if let Some(fm) = self.project.forms.get_mut(idx) {
                        if let Some(ctrl) = fm.form.controls.iter_mut().find(|c| c.name == ctrl_name) {
                            let db_type = ctrl.properties.get_string("DbType").unwrap_or("SQLite").to_string();
                            let conn = match db_type.as_str() {
                                "SQLite" => {
                                    let path = ctrl.properties.get_string("DbPath").unwrap_or("database.db");
                                    format!("Data Source={}", path)
                                }
                                "PostgreSQL" => {
                                    let host = ctrl.properties.get_string("DbHost").unwrap_or("localhost");
                                    let port = ctrl.properties.get_string("DbPort").unwrap_or("5432");
                                    let db = ctrl.properties.get_string("DbName").unwrap_or("");
                                    let user = ctrl.properties.get_string("DbUser").unwrap_or("postgres");
                                    let pass = ctrl.properties.get_string("DbPassword").unwrap_or("");
                                    format!("Host={};Port={};Database={};Username={};Password={}", host, port, db, user, pass)
                                }
                                "MySQL" => {
                                    let host = ctrl.properties.get_string("DbHost").unwrap_or("localhost");
                                    let port = ctrl.properties.get_string("DbPort").unwrap_or("3306");
                                    let db = ctrl.properties.get_string("DbName").unwrap_or("");
                                    let user = ctrl.properties.get_string("DbUser").unwrap_or("root");
                                    let pass = ctrl.properties.get_string("DbPassword").unwrap_or("");
                                    format!("Server={};Port={};Database={};Uid={};Pwd={}", host, port, db, user, pass)
                                }
                                _ => String::new(),
                            };
                            ctrl.properties.set("ConnectionString", conn);
                            self.status_bar.message = "Connection string built".to_string();
                        }
                    }
                }
            }
            "test_conn" => {
                // Test connection — read connection string and attempt to connect
                let conn_str = if let Some(idx) = self.current_form_idx() {
                    self.project.forms.get(idx).and_then(|fm| {
                        fm.form.controls.iter().find(|c| c.name == ctrl_name)
                            .and_then(|c| c.properties.get_string("ConnectionString"))
                            .map(|s| s.to_string())
                    }).unwrap_or_default()
                } else {
                    String::new()
                };

                if conn_str.is_empty() {
                    self.properties.conn_status = "⚠ No connection string".to_string();
                    self.properties.conn_tables.clear();
                } else {
                    // Try to connect via vybe_host if available
                    #[cfg(feature = "database")]
                    {
                        match vybe_host::test_connection_and_list_tables(&conn_str) {
                            Ok(tables) => {
                                self.properties.conn_status = format!("✓ Connected — {} tables", tables.len());
                                self.properties.conn_tables = tables;
                            }
                            Err(e) => {
                                self.properties.conn_status = format!("✗ {}", e);
                                self.properties.conn_tables.clear();
                            }
                        }
                    }
                    #[cfg(not(feature = "database"))]
                    {
                        self.properties.conn_status = format!("✓ Connection string set: {}", &conn_str[..conn_str.len().min(40)]);
                        self.properties.conn_tables.clear();
                    }
                }
                self.status_bar.message = self.properties.conn_status.clone();
            }
            _ => {}
        }
    }
}
