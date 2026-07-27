use super::{App, EditAction, PaletteAction, ProjectSearchHit, Tab, TabContent};
use crate::editor::Editor as MyEditor;
use crate::language::load_language;
use crate::lsp_client::LspRequest;
use vybe_widgets::code_editor_widget::CodeEditorWidget;

/// Insert `snippet` into a VB.NET class body just before its `End Class`.
/// Mirrors the legacy designer — if there's no `End Class`, appends the
/// snippet at the end.
fn insert_before_end_class(code: &str, snippet: &str) -> String {
    let lower = code.to_lowercase();
    if let Some(idx) = lower.rfind("end class") {
        let (head, tail) = code.split_at(idx);
        format!("{}\n\n{}\n{}", head.trim_end(), snippet, tail)
    } else {
        format!("{}\n\n{}", code, snippet)
    }
}

/// Line (0-indexed) of the first line that contains `needle`. Case-sensitive.
fn locate_substring_line(code: &str, needle: &str) -> Option<usize> {
    code.lines()
        .enumerate()
        .find_map(|(i, ln)| if ln.contains(needle) { Some(i) } else { None })
}

/// Find an existing VB.NET event handler. Returns the 0-indexed line of the
/// `Sub` declaration (the body starts on the next line).
///
/// Recognizes three patterns (all case-insensitive):
///   1. `Private Sub {handler_name}(…)`         — legacy naming convention
///   2. `Public Sub {handler_name}(…)`
///   3. Any `Sub X(…)` whose `Handles` clause mentions `{target}.{event}` or `Me.{event}`
fn find_handler_line(
    code: &str,
    handler_name: &str,
    target: &str,
    event: &str,
    is_form: bool,
) -> Option<usize> {
    let hn = handler_name.to_lowercase();
    let target_dot = format!("{}.{}", target.to_lowercase(), event.to_lowercase());
    let me_dot = format!("me.{}", event.to_lowercase());

    for (i, line) in code.lines().enumerate() {
        let lower = line.trim_start().to_lowercase();
        // Name match: Private/Public/Friend/Protected Sub {hn}(...)
        // or bare "Sub {hn}(".
        if let Some(sub_start) = lower.find("sub ") {
            let after = &lower[sub_start + 4..];
            let end = after.find('(').unwrap_or(after.len());
            let name_part = after[..end].trim();
            if name_part == hn {
                return Some(i);
            }
        }
        // Handles-clause match.
        if lower.contains(" handles ") {
            if lower.contains(&target_dot) || (is_form && lower.contains(&me_dot)) {
                return Some(i);
            }
        }
    }
    None
}

const OUTPUT_MAX_LINES: usize = 5_000;

impl App {
    /// Push a line to the output panel and auto-show it. Maintains a
    /// rolling buffer of at most `OUTPUT_MAX_LINES` (drops the oldest).
    pub(crate) fn output_push(&mut self, line: String) {
        self.output_lines_buffer.push(line);
        if self.output_lines_buffer.len() > OUTPUT_MAX_LINES {
            let drop = self.output_lines_buffer.len() - OUTPUT_MAX_LINES;
            self.output_lines_buffer.drain(..drop);
        }
        self.output_panel
            .set_output_lines(&self.output_lines_buffer);
        if !self.output_panel.visible() {
            self.output_panel.set_visible(true);
            self.relayout();
        }
        self.output_panel.scroll_to_bottom();
    }

    pub(super) fn flush_code_to_project(&mut self) {
        for tab in &self.tabs {
            if let TabContent::Code(cw) = &tab.content {
                let code = cw.my_editor.rope.to_string();
                let tab_name = &tab.name;
                if let Some(form_name) = tab_name.strip_suffix(".vb") {
                    if let Some(fm) = self
                        .project
                        .forms
                        .iter_mut()
                        .find(|fm| fm.form.name == form_name)
                    {
                        fm.set_user_code(code);
                        continue;
                    }
                }
                if let Some(cf) = self
                    .project
                    .code_files
                    .iter_mut()
                    .find(|cf| cf.name == *tab_name)
                {
                    cf.code = code;
                }
            }
            if let TabContent::Form(f) = &tab.content {
                if let Some(fm) = self
                    .project
                    .forms
                    .iter_mut()
                    .find(|fm| fm.form.name == f.form.name)
                {
                    fm.form = f.form.clone();
                }
            }
        }
    }

    pub(super) fn save_project(&mut self) {
        self.flush_code_to_project();
        self.sync_active_form_to_project();
        if let Some(path) = self.project_path.clone() {
            match vybe_platform_dotnet::winforms::designer::serialization::save_project_auto(&self.project, &path) {
                Ok(_) => self.output_push(format!("Saved: {}", path)),
                Err(e) => self.output_push(format!("Save error: {}", e)),
            }
        } else {
            if let Some(path) = rfd::FileDialog::new()
                .add_filter("VB Project", &["vbproj"])
                .save_file()
            {
                let path_str = path.to_string_lossy().to_string();
                self.project_path = Some(path_str.clone());
                match vybe_platform_dotnet::winforms::designer::serialization::save_project_auto(
                    &self.project,
                    &path_str,
                ) {
                    Ok(_) => self.output_push(format!("Saved: {}", path_str)),
                    Err(e) => self.output_push(format!("Save error: {}", e)),
                }
            }
        }
    }

    pub(super) fn run_project(&mut self) {
        self.flush_code_to_project();
        self.sync_active_form_to_project();
        let path = match self.project_path.clone() {
            Some(p) => p,
            None => {
                self.output_push("Save the project first.".to_string());
                return;
            }
        };
        let _ = vybe_platform_dotnet::winforms::designer::serialization::save_project_auto(&self.project, &path);
        let config_flag = match self.build_config {
            super::BuildConfig::Debug => "--debug",
            super::BuildConfig::Release => "--release",
        };
        let vybec = std::env::current_exe()
            .ok()
            .and_then(|p| p.parent().map(|d| d.join("vybex")))
            .unwrap_or_else(|| std::path::PathBuf::from("vybex"));
        self.output_push(format!(
            "Building ({})...",
            if self.build_config == super::BuildConfig::Debug {
                "Debug"
            } else {
                "Release"
            }
        ));
        match std::process::Command::new(&vybec)
            .arg(&path)
            .arg(config_flag)
            .stdout(std::process::Stdio::piped())
            .stderr(std::process::Stdio::piped())
            .spawn()
        {
            Ok(mut child) => {
                if let Some(stdout) = child.stdout.take() {
                    use std::io::BufRead;
                    let reader = std::io::BufReader::new(stdout);
                    for line in reader.lines() {
                        if let Ok(l) = line {
                            self.output_push(l);
                        }
                    }
                }
                if let Some(stderr) = child.stderr.take() {
                    use std::io::BufRead;
                    let reader = std::io::BufReader::new(stderr);
                    for line in reader.lines() {
                        if let Ok(l) = line {
                            self.output_push(format!("ERR: {}", l));
                        }
                    }
                }
                self.run_child = Some(child);
                self.output_push(format!("Running project: {}", path));
            }
            Err(e) => self.output_push(format!("Could not launch vybec: {}", e)),
        }
    }

    pub(super) fn stop_project(&mut self) {
        if let Some(ref mut child) = self.run_child {
            let _ = child.kill();
            self.output_push("Stopped.".to_string());
        }
        self.run_child = None;
    }

    pub(super) fn add_existing_form(&mut self) {
        let Some(paths) = rfd::FileDialog::new()
            .set_title("Add Existing Form")
            .add_filter("VB Forms", &["vb"])
            .pick_files()
        else {
            return;
        };
        for path in paths {
            match vybe_platform_dotnet::winforms::designer::load_form_vb(&path) {
                Ok(fm) => {
                    let name = fm.form.name.clone();
                    if self.project.forms.iter().all(|f| f.form.name != name) {
                        let form_clone = fm.form.clone();
                        self.project.forms.push(fm);
                        if let Some(idx) = self
                            .tabs
                            .iter()
                            .position(|t| matches!(&t.content, TabContent::Form(_)))
                        {
                            if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                                fd.form = form_clone;
                                fd.selected_controls.clear();
                            }
                            self.active_tab = idx;
                        }
                    }
                }
                Err(e) => self.output_push(format!("Failed to load form: {}", e)),
            }
        }
    }

    /// Insert a VB.NET event handler for `(target, event)` into the form's
    /// code-behind (user_code) if it doesn't already exist, then open /
    /// focus a code tab showing it AND jump the cursor to the handler's body.
    /// `target` is either a control name or the form name; `is_form` selects
    /// `Handles Me.X` vs `Handles X.Y`.
    pub(super) fn open_or_generate_event_handler(
        &mut self,
        form_name: &str,
        target: &str,
        event: &str,
        is_form: bool,
    ) {
        let params = vybe_platform_dotnet::winforms::designer::EventType::from_name(event)
            .map(|et| et.parameters().to_string())
            .unwrap_or_else(|| "sender As Object, e As EventArgs".to_string());
        let handler_name = format!("{}_{}", target, event);
        let handles = if is_form {
            format!("Handles Me.{}", event)
        } else {
            format!("Handles {}.{}", target, event)
        };
        let sub_decl = format!("Private Sub {}({}) {}", handler_name, params, handles);
        let sub_body = format!("{}\n    ' TODO: Add your code here\nEnd Sub", sub_decl);

        let form_module = self
            .project
            .forms
            .iter_mut()
            .find(|f| f.form.name == form_name);
        let Some(fm) = form_module else {
            return;
        };
        let current = fm.get_user_code().to_string();

        // Detect: any `Private Sub {handler_name}(` OR `Sub {handler_name}(`
        // OR an existing Sub with `Handles {target or Me}.{event}`.
        let existing_line = find_handler_line(&current, &handler_name, target, event, is_form);
        let (new_code, jump_line) = match existing_line {
            Some(line) => (current, line),
            None => {
                let new_code = insert_before_end_class(&current, &sub_body);
                // The body line is sub_decl's line + 1.
                let decl_line = locate_substring_line(&new_code, &sub_decl).unwrap_or(0);
                (new_code, decl_line.saturating_add(1))
            }
        };
        if new_code != fm.get_user_code() {
            fm.set_user_code(new_code.clone());
        }

        // Find or create the code tab for this form's code-behind.
        let tab_name = format!("{}.vb", form_name);
        let existing = self
            .tabs
            .iter()
            .position(|t| t.name == tab_name && matches!(&t.content, TabContent::Code(_)));
        if let Some(idx) = existing {
            if let TabContent::Code(w) = &mut self.tabs[idx].content {
                w.set_buffer_text(&mut self.font_system, &new_code);
                w.set_cursor_pos(jump_line, 4);
            }
            self.active_tab = idx;
        } else {
            let lang = load_language("vb")
                .or_else(|| load_language("rust"))
                .expect("language not found");
            let my_editor = MyEditor::from_text(&new_code, &lang);
            let mut widget = CodeEditorWidget::new(my_editor.inner, &mut self.font_system);
            widget.set_cursor_pos(jump_line, 4);
            self.tabs.push(Tab {
                name: tab_name,
                path: None,
                content: TabContent::Code(widget),
                is_sticky: true,
                buffer: None,
                is_modified: false,
            });
            self.active_tab = self.tabs.len() - 1;
        }
        self.needs_redraw = true;
    }

    /// Copy the active form-designer's `Form` back into the matching
    /// `FormModule` in the project, then regenerate designer code. Call this
    /// after any designer mutation so switching forms or saving picks up
    /// the latest edits.
    pub(crate) fn sync_active_form_to_project(&mut self) {
        let Some(tab) = self.tabs.get(self.active_tab) else {
            return;
        };
        let TabContent::Form(f) = &tab.content else {
            return;
        };
        let name = f.form.name.clone();
        let form_clone = f.form.clone();
        if let Some(fm) = self
            .project
            .forms
            .iter_mut()
            .find(|fm| fm.form.name == name)
        {
            fm.form = form_clone;
            fm.sync_designer_code();
        }
    }

    pub(super) fn add_existing_code(&mut self) {
        let Some(paths) = rfd::FileDialog::new()
            .set_title("Add Existing Code File")
            .add_filter("Code Files", &["vb", "bas"])
            .pick_files()
        else {
            return;
        };
        for path in paths {
            let name = path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string();
            let code = match vybe_compiler::projects::read_text_file(&path) {
                Ok(c) => c,
                Err(e) => {
                    self.output_push(format!("Cannot read file: {}", e));
                    continue;
                }
            };
            if self.project.code_files.iter().all(|cf| cf.name != name) {
                self.project
                    .code_files
                    .push(vybe_platform_dotnet::winforms::designer::project::CodeFile {
                        name: name.clone(),
                        code: code.clone(),
                    });
                let lang = load_language("vb")
                    .or_else(|| load_language("rust"))
                    .expect("language not found");
                let my_editor = MyEditor::from_text(&code, &lang);
                let uri = format!("file:///project/{}", name);
                let widget = {
                    let text = my_editor.rope.to_string();
                    self.lsp.send(LspRequest::Init(text, "vb".to_string(), uri));
                    CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
                };
                self.tabs.push(Tab {
                    name,
                    path: None,
                    content: TabContent::Code(widget),
                    is_sticky: true,
                    buffer: None,
                    is_modified: false,
                });
                self.active_tab = self.tabs.len() - 1;
            }
        }
    }

    pub(super) fn remove_project_item(&mut self, name: &str) {
        let removed = self.project.remove_form(name) || self.project.remove_code_file(name);
        if removed {
            let tab_name_vb = format!("{}.vb", name);
            self.tabs
                .retain(|t| t.name != name && t.name != tab_name_vb);
            if self.active_tab >= self.tabs.len() && !self.tabs.is_empty() {
                self.active_tab = self.tabs.len() - 1;
            }
            if let Some(idx) = self
                .tabs
                .iter()
                .position(|t| matches!(&t.content, TabContent::Form(_)))
            {
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

    pub(super) fn dispatch_edit_action(&mut self, action: EditAction) {
        if self.active_tab >= self.tabs.len() {
            return;
        }
        match &mut self.tabs[self.active_tab].content {
            TabContent::Form(f) => match action {
                EditAction::Undo => {
                    f.undo();
                }
                EditAction::Redo => {
                    f.redo();
                }
                EditAction::Delete => {
                    f.push_undo_snapshot();
                    let sel = f.selected_controls.clone();
                    f.form.controls.retain(|c| !sel.contains(&c.id));
                    f.selected_controls.clear();
                }
                EditAction::Cut => {
                    f.push_undo_snapshot();
                    self.control_clipboard = f
                        .selected_controls
                        .iter()
                        .filter_map(|id| f.form.controls.iter().find(|c| c.id == *id).cloned())
                        .collect();
                    let sel = f.selected_controls.clone();
                    f.form.controls.retain(|c| !sel.contains(&c.id));
                    f.selected_controls.clear();
                }
                EditAction::Copy => {
                    self.control_clipboard = f
                        .selected_controls
                        .iter()
                        .filter_map(|id| f.form.controls.iter().find(|c| c.id == *id).cloned())
                        .collect();
                }
                EditAction::Paste => {
                    f.push_undo_snapshot();
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
                                if let Ok(n) = c.name[base.len()..].parse::<u32>() {
                                    max = max.max(n);
                                }
                            }
                        }
                        ctrl.name = format!("{}{}", base, max + 1);
                        new_ids.push(ctrl.id);
                        f.form.controls.push(ctrl);
                    }
                    f.selected_controls = new_ids;
                }
            },
            TabContent::Code(cw) => {
                match action {
                    EditAction::Undo => {
                        let (cl, ci) = cw.cursor_pos();
                        if let Some((text, line, col)) = cw.my_editor.undo(cl, ci) {
                            cw.set_buffer_text(&mut self.font_system, &text);
                            cw.set_cursor_pos(line, col);
                        }
                    }
                    EditAction::Redo => {
                        let (cl, ci) = cw.cursor_pos();
                        if let Some((text, line, col)) = cw.my_editor.redo(cl, ci) {
                            cw.set_buffer_text(&mut self.font_system, &text);
                            cw.set_cursor_pos(line, col);
                        }
                    }
                    EditAction::Cut => {
                        if let Some(t) = cw.copy_selection_text() {
                            cw.my_editor
                                .save_snapshot(cw.cursor_pos().0, cw.cursor_pos().1);
                            if let Some(cb) = &mut self.clipboard {
                                let _ = cb.set_text(t);
                            }
                            cw.action_delete(&mut self.font_system);
                        }
                    }
                    EditAction::Copy => {
                        if let Some(t) = cw.copy_selection_text() {
                            if let Some(cb) = &mut self.clipboard {
                                let _ = cb.set_text(t);
                            }
                        }
                    }
                    EditAction::Paste => {
                        if let Some(cb) = &mut self.clipboard {
                            if let Ok(t) = cb.get_text() {
                                cw.my_editor
                                    .save_snapshot(cw.cursor_pos().0, cw.cursor_pos().1);
                                let byte_off =
                                    cw.compute_byte_offset(cw.cursor_pos().0, cw.cursor_pos().1);
                                let (new_line, new_col) =
                                    cw.my_editor.insert_string(byte_off, &t, &cw.lang_def);
                                cw.set_buffer_text(
                                    &mut self.font_system,
                                    &cw.my_editor.rope().to_string(),
                                );
                                cw.set_cursor_pos(new_line, new_col);
                            }
                        }
                    }
                    EditAction::Delete => {
                        cw.my_editor
                            .save_snapshot(cw.cursor_pos().0, cw.cursor_pos().1);
                        cw.action_delete(&mut self.font_system);
                    }
                }
                cw.needs_reshape = true;
                cw.sync();
            }
            TabContent::Resources(_) => {}
        }
    }

    pub(super) fn create_resource_editor_from_project(
        project: &vybe_platform_dotnet::winforms::designer::project::Project,
    ) -> vybe_widgets::ResourceEditor {
        let mut editor = vybe_widgets::ResourceEditor::new();
        for rm in &project.resource_files {
            for item in &rm.resources {
                editor.entries.push(vybe_widgets::ResourceEntry {
                    name: item.name.clone(),
                    value: item.value.clone(),
                    comment: item.comment.clone().unwrap_or_default(),
                    tab: match item.resource_type {
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::String => {
                            vybe_widgets::ResourceTab::Strings
                        }
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::Image => {
                            vybe_widgets::ResourceTab::Images
                        }
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::Icon => {
                            vybe_widgets::ResourceTab::Icons
                        }
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::Audio => {
                            vybe_widgets::ResourceTab::Audio
                        }
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::File => {
                            vybe_widgets::ResourceTab::Files
                        }
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::Other => {
                            vybe_widgets::ResourceTab::Other
                        }
                    },
                    file_name: item.file_name.clone(),
                });
            }
        }
        editor
    }

    pub(super) fn process_resource_event(
        evt: vybe_widgets::ResourceEditorEvent,
        r: &mut vybe_widgets::ResourceEditor,
        project: &mut vybe_platform_dotnet::winforms::designer::project::Project,
    ) {
        match evt {
            vybe_widgets::ResourceEditorEvent::AddResource(tab) => {
                if tab.is_file_based() {
                    let mut dialog = rfd::FileDialog::new();
                    let (filter_name, exts): (&str, Vec<&str>) = match tab {
                        vybe_widgets::ResourceTab::Images => (
                            "Images",
                            vec!["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"],
                        ),
                        vybe_widgets::ResourceTab::Icons => ("Icons", vec!["ico"]),
                        vybe_widgets::ResourceTab::Audio => {
                            ("Audio", vec!["wav", "mp3", "ogg", "flac", "aiff"])
                        }
                        _ => ("All Files", vec!["*"]),
                    };
                    if !exts.is_empty() && exts[0] != "*" {
                        dialog = dialog.add_filter(filter_name, &exts);
                    }
                    dialog = dialog.add_filter("All Files", &["*"]);
                    if let Some(paths) = dialog.pick_files() {
                        for p in paths {
                            let path_str = p.to_string_lossy().to_string();
                            let name = p
                                .file_stem()
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
                            let rt = match res_tab {
                                vybe_widgets::ResourceTab::Images => {
                                    vybe_platform_dotnet::winforms::designer::resources::ResourceType::Image
                                }
                                vybe_widgets::ResourceTab::Icons => {
                                    vybe_platform_dotnet::winforms::designer::resources::ResourceType::Icon
                                }
                                vybe_widgets::ResourceTab::Audio => {
                                    vybe_platform_dotnet::winforms::designer::resources::ResourceType::Audio
                                }
                                vybe_widgets::ResourceTab::Files => {
                                    vybe_platform_dotnet::winforms::designer::resources::ResourceType::File
                                }
                                _ => vybe_platform_dotnet::winforms::designer::resources::ResourceType::String,
                            };
                            if project.resource_files.is_empty() {
                                project
                                    .resource_files
                                    .push(vybe_platform_dotnet::winforms::designer::ResourceManager::new());
                            }
                            if let Some(rm) = project.resource_files.first_mut() {
                                rm.resources
                                    .push(vybe_platform_dotnet::winforms::designer::ResourceItem::new_file(
                                        name, &path_str, rt,
                                    ));
                            }
                        }
                    }
                } else {
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
                    Self::sync_resources_to_project(r, project);
                }
            }
            vybe_widgets::ResourceEditorEvent::BrowseFile(idx) => {
                if idx < r.entries.len() {
                    let entry = &r.entries[idx];
                    let rt = &entry.tab;
                    let mut dialog = rfd::FileDialog::new();
                    let exts: Vec<&str> = match rt {
                        vybe_widgets::ResourceTab::Images => {
                            vec!["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"]
                        }
                        vybe_widgets::ResourceTab::Icons => vec!["ico"],
                        vybe_widgets::ResourceTab::Audio => {
                            vec!["wav", "mp3", "ogg", "flac", "aiff"]
                        }
                        _ => vec![],
                    };
                    if !exts.is_empty() {
                        dialog = dialog.add_filter("Supported", &exts);
                    }
                    dialog = dialog.add_filter("All Files", &["*"]);
                    if let Some(path) = dialog.pick_file() {
                        let path_str = path.to_string_lossy().to_string();
                        let name = path
                            .file_stem()
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
                if project.resource_files.is_empty() {
                    project
                        .resource_files
                        .push(vybe_platform_dotnet::winforms::designer::ResourceManager::new());
                }
                if let Some(rm) = project.resource_files.first_mut() {
                    let mut item = vybe_platform_dotnet::winforms::designer::ResourceItem::new_string(name, value);
                    item.resource_type = match tab {
                        vybe_widgets::ResourceTab::Other => {
                            vybe_platform_dotnet::winforms::designer::resources::ResourceType::Other
                        }
                        _ => vybe_platform_dotnet::winforms::designer::resources::ResourceType::String,
                    };
                    item.comment = if comment.is_empty() {
                        None
                    } else {
                        Some(comment)
                    };
                    rm.resources.push(item);
                }
            }
            vybe_widgets::ResourceEditorEvent::EditCommitted(_, _, _) => {
                Self::sync_resources_to_project(r, project);
            }
            _ => {}
        }
    }

    pub(super) fn sync_resources_to_project(
        r: &vybe_widgets::ResourceEditor,
        project: &mut vybe_platform_dotnet::winforms::designer::project::Project,
    ) {
        if project.resource_files.is_empty() {
            project
                .resource_files
                .push(vybe_platform_dotnet::winforms::designer::ResourceManager::new());
        }
        if let Some(rm) = project.resource_files.first_mut() {
            rm.resources.clear();
            for entry in &r.entries {
                let rt = match entry.tab {
                    vybe_widgets::ResourceTab::Strings => {
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::String
                    }
                    vybe_widgets::ResourceTab::Images => {
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::Image
                    }
                    vybe_widgets::ResourceTab::Icons => {
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::Icon
                    }
                    vybe_widgets::ResourceTab::Audio => {
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::Audio
                    }
                    vybe_widgets::ResourceTab::Files => {
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::File
                    }
                    vybe_widgets::ResourceTab::Other => {
                        vybe_platform_dotnet::winforms::designer::resources::ResourceType::Other
                    }
                };
                if entry.tab.is_file_based() {
                    let mut item = vybe_platform_dotnet::winforms::designer::ResourceItem::new_file(
                        entry.name.clone(),
                        &entry.value,
                        rt,
                    );
                    item.comment = if entry.comment.is_empty() {
                        None
                    } else {
                        Some(entry.comment.clone())
                    };
                    rm.resources.push(item);
                } else {
                    let mut item = vybe_platform_dotnet::winforms::designer::ResourceItem::new_string(
                        entry.name.clone(),
                        entry.value.clone(),
                    );
                    item.resource_type = rt;
                    item.comment = if entry.comment.is_empty() {
                        None
                    } else {
                        Some(entry.comment.clone())
                    };
                    rm.resources.push(item);
                }
            }
        }
    }

    /// Send a completion request to the LSP for the current cursor position.
    pub(super) fn trigger_completion(&self) {
        if let Some(tab) = self.tabs.get(self.active_tab) {
            if let TabContent::Code(w) = &tab.content {
                let (line, col) = w.cursor_pos();
                let uri = App::tab_uri(tab);
                self.lsp
                    .send(LspRequest::Completion(uri, line as u32, col as u32));
            }
        }
    }

    /// Execute an action chosen from the command palette.
    pub(super) fn execute_palette_command(&mut self, action: PaletteAction) {
        use crate::form_designer_tab::MenuAction as M;
        use PaletteAction::*;
        match action {
            Menu(M::NewProject)
            | Menu(M::OpenProject)
            | Menu(M::AddExistingForm)
            | Menu(M::AddExistingCode)
            | Menu(M::ProjectProperties) => {
                // These open file-pickers / dialogs. Easiest to route via the
                // existing menu click path — emulate by calling the same
                // command methods directly.
                match action {
                    Menu(M::OpenProject) => self.palette_open_project(),
                    Menu(M::AddExistingForm) => self.add_existing_form(),
                    Menu(M::AddExistingCode) => self.add_existing_code(),
                    Menu(M::ProjectProperties) => {
                        self.project_props_dialog.open(&self.project);
                    }
                    _ => {}
                }
            }
            Menu(M::SaveProject) => self.save_project(),
            Menu(M::SaveAs) => self.palette_save_project_as(),
            Menu(M::Exit) => {
                std::process::exit(0);
            }
            Menu(M::AddForm) => self.palette_add_form(),
            Menu(M::AddModule) => self.palette_add_module(),
            Menu(M::AddResourceFile) => self.palette_add_resource_file(),
            Menu(M::RunProject) => self.run_project(),
            Menu(M::StopProject) => self.stop_project(),
            Menu(M::Undo) => self.dispatch_edit_action(EditAction::Undo),
            Menu(M::Redo) => self.dispatch_edit_action(EditAction::Redo),
            Menu(M::Cut) => self.dispatch_edit_action(EditAction::Cut),
            Menu(M::Copy) => self.dispatch_edit_action(EditAction::Copy),
            Menu(M::Paste) => self.dispatch_edit_action(EditAction::Paste),
            Menu(M::Delete) => self.dispatch_edit_action(EditAction::Delete),

            Edit(a) => self.dispatch_edit_action(a),
            ToggleOutput => {
                let v = self.output_panel.visible();
                self.output_panel.set_visible(!v);
            }
            ToggleProblems => {
                self.output_panel.set_visible(true);
                self.output_panel
                    .set_active_tab(vybe_widgets::output_panel::OutputTab::Problems);
            }
            CloseTab => {
                self.close_tab(self.active_tab);
            }
            CloseOthers => {
                self.close_others(self.active_tab);
            }
            CloseAll => {
                self.close_all_tabs();
            }
            NextTab => {
                if !self.tabs.is_empty() {
                    self.active_tab = (self.active_tab + 1) % self.tabs.len();
                }
            }
            PrevTab => {
                if !self.tabs.is_empty() {
                    self.active_tab = if self.active_tab == 0 {
                        self.tabs.len() - 1
                    } else {
                        self.active_tab - 1
                    };
                }
            }
            FindInFile => {
                if let Some(TabContent::Code(cw)) =
                    self.tabs.get_mut(self.active_tab).map(|t| &mut t.content)
                {
                    cw.is_search_open = true;
                }
            }
            FindInProject => {
                self.is_project_search = true;
                self.project_search_query.clear();
                self.project_search_results.clear();
                self.project_search_selected = 0;
            }
            GoToLine => {
                self.goto_line_open = true;
                self.goto_line_query.clear();
            }
        }
        self.needs_redraw = true;
    }

    /// Close the tab at `idx`. Keeps at least the Form-Designer tab if
    /// that's the one being closed (designer is always pinned).
    pub(super) fn close_tab(&mut self, idx: usize) {
        if idx >= self.tabs.len() {
            return;
        }
        if self.tabs[idx].is_sticky {
            return;
        }
        self.tabs.remove(idx);
        if self.active_tab >= self.tabs.len() && !self.tabs.is_empty() {
            self.active_tab = self.tabs.len() - 1;
        }
    }

    /// Close everything except the keep_idx tab (and the always-pinned ones).
    pub(super) fn close_others(&mut self, keep_idx: usize) {
        let keep_uid: Option<String> = self.tabs.get(keep_idx).map(|t| t.name.clone());
        self.tabs
            .retain(|t| t.is_sticky || Some(t.name.clone()) == keep_uid);
        self.active_tab = self
            .tabs
            .iter()
            .position(|t| Some(t.name.clone()) == keep_uid)
            .unwrap_or(0);
    }

    pub(super) fn close_all_tabs(&mut self) {
        self.tabs.retain(|t| t.is_sticky);
        if self.active_tab >= self.tabs.len() && !self.tabs.is_empty() {
            self.active_tab = self.tabs.len() - 1;
        }
    }

    // ── Palette-action helpers (these wrap flows previously inlined in
    // app_mouse.rs menu handling). ────────────────────────────────────────

    fn palette_open_project(&mut self) {
        if let Some(path) = rfd::FileDialog::new()
            .add_filter("VB Project", &["vbproj", "vbp"])
            .pick_file()
        {
            let path_str = path.to_string_lossy().to_string();
            if let Ok(proj) = vybe_platform_dotnet::winforms::designer::serialization::load_project_auto(&path_str) {
                self.tabs.retain(|t| {
                    !matches!(&t.content, TabContent::Code(_))
                        && !matches!(&t.content, TabContent::Resources(_))
                });
                if let Some(first_form) = proj.forms.first() {
                    if let Some(idx) = self
                        .tabs
                        .iter()
                        .position(|t| matches!(&t.content, TabContent::Form(_)))
                    {
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
        }
    }

    fn palette_save_project_as(&mut self) {
        if let Some(path) = rfd::FileDialog::new()
            .add_filter("VB Project", &["vbproj"])
            .save_file()
        {
            let path_str = path.to_string_lossy().to_string();
            self.project_path = Some(path_str.clone());
            let _ =
                vybe_platform_dotnet::winforms::designer::serialization::save_project_auto(&self.project, &path_str);
        }
    }

    fn palette_add_form(&mut self) {
        self.sync_active_form_to_project();
        let name = format!("Form{}", self.project.forms.len() + 1);
        let mut form = vybe_platform_dotnet::winforms::designer::Form::new(name.clone());
        form.width = 640;
        form.height = 480;
        let form_clone = form.clone();
        let fm = vybe_platform_dotnet::winforms::designer::project::FormModule::new_classic(form);
        self.project.forms.push(fm);
        if let Some(idx) = self
            .tabs
            .iter()
            .position(|t| matches!(&t.content, TabContent::Form(_)))
        {
            if let TabContent::Form(fd) = &mut self.tabs[idx].content {
                fd.form = form_clone;
                fd.selected_controls.clear();
            }
            self.active_tab = idx;
        }
    }

    fn palette_add_module(&mut self) {
        let name = format!("Module{}.vb", self.project.code_files.len() + 1);
        self.project
            .code_files
            .push(vybe_platform_dotnet::winforms::designer::project::CodeFile {
                name: name.clone(),
                code: format!("Module {}\n\nEnd Module\n", name.replace(".vb", "")),
            });
    }

    fn palette_add_resource_file(&mut self) {
        if self.project.resource_files.is_empty() {
            self.project
                .resource_files
                .push(vybe_platform_dotnet::winforms::designer::ResourceManager::new());
        }
        if let Some(idx) = self
            .tabs
            .iter()
            .position(|t| matches!(&t.content, TabContent::Resources(_)))
        {
            self.active_tab = idx;
        } else {
            let editor = Self::create_resource_editor_from_project(&self.project);
            self.tabs.push(Tab {
                name: "Resources".to_string(),
                path: None,
                content: TabContent::Resources(editor),
                is_sticky: true,
                buffer: None,
                is_modified: false,
            });
            self.active_tab = self.tabs.len() - 1;
        }
    }

    // ── Project-wide search ───────────────────────────────────────────────

    /// Re-run search with the current `project_search_query`, updating the
    /// results list. Search covers every code file and form's user_code.
    /// Max 500 hits to keep the UI snappy.
    pub(super) fn run_project_search(&mut self) {
        self.project_search_results.clear();
        self.project_search_selected = 0;
        let q = self.project_search_query.trim();
        if q.len() < 2 {
            return;
        }
        let needle = q.to_lowercase();
        let mut hits: Vec<ProjectSearchHit> = Vec::new();

        fn scan_into(
            out: &mut Vec<ProjectSearchHit>,
            needle: &str,
            file: &str,
            code: &str,
            cap: usize,
        ) {
            for (li, line) in code.lines().enumerate() {
                if out.len() >= cap {
                    return;
                }
                if line.to_lowercase().contains(needle) {
                    let snippet = line.trim().chars().take(120).collect::<String>();
                    out.push(ProjectSearchHit {
                        file: file.to_string(),
                        line: li,
                        snippet,
                    });
                }
            }
        }

        const CAP: usize = 500;
        for fm in &self.project.forms {
            if hits.len() >= CAP {
                break;
            }
            scan_into(
                &mut hits,
                &needle,
                &format!("{}.vb", fm.form.name),
                fm.get_user_code(),
                CAP,
            );
        }
        for cf in &self.project.code_files {
            if hits.len() >= CAP {
                break;
            }
            scan_into(&mut hits, &needle, &cf.name, &cf.code, CAP);
        }
        for tab in &self.tabs {
            if let TabContent::Code(cw) = &tab.content {
                if hits.len() >= CAP {
                    break;
                }
                let code = cw.my_editor.rope.to_string();
                scan_into(&mut hits, &needle, &tab.name, &code, CAP);
            }
        }

        // Also walk the filesystem from the project directory
        if hits.len() < CAP {
            if let Some(proj_path) = &self.project_path {
                if let Some(root) = std::path::Path::new(proj_path).parent() {
                    search_dir(root, &needle, &mut hits, CAP);
                }
            }
        }

        self.project_search_results = hits;
    }

    /// Open a search hit in a code tab.
    pub(super) fn project_search_open_hit(&mut self, hit: &ProjectSearchHit) {
        // Does a tab already display this file?
        if let Some(idx) = self
            .tabs
            .iter()
            .position(|t| t.name == hit.file && matches!(&t.content, TabContent::Code(_)))
        {
            self.active_tab = idx;
            if let TabContent::Code(cw) = &mut self.tabs[idx].content {
                cw.set_cursor_pos(hit.line, 0);
            }
            return;
        }
        // Otherwise load from project.
        let code = self
            .project
            .forms
            .iter()
            .find(|fm| format!("{}.vb", fm.form.name) == hit.file)
            .map(|fm| fm.get_user_code().to_string())
            .or_else(|| {
                self.project
                    .code_files
                    .iter()
                    .find(|cf| cf.name == hit.file)
                    .map(|cf| cf.code.clone())
            });
        let Some(code) = code else {
            return;
        };

        let lang = load_language("vb")
            .or_else(|| load_language("rust"))
            .expect("language not found");
        let my_editor = MyEditor::from_text(&code, &lang);
        let mut widget = CodeEditorWidget::new(my_editor.inner, &mut self.font_system);
        widget.set_cursor_pos(hit.line, 0);
        self.tabs.push(Tab {
            name: hit.file.clone(),
            path: None,
            content: TabContent::Code(widget),
            is_sticky: false,
            buffer: None,
            is_modified: false,
        });
        self.active_tab = self.tabs.len() - 1;
    }
}

// ── Filesystem search helpers ──────────────────────────────────────────────

fn is_text_extension(ext: &str) -> bool {
    matches!(
        ext,
        "rs" | "vb"
            | "bas"
            | "frm"
            | "cls"
            | "cs"
            | "js"
            | "mjs"
            | "ts"
            | "tsx"
            | "py"
            | "php"
            | "rb"
            | "dart"
            | "java"
            | "go"
            | "c"
            | "h"
            | "cpp"
            | "cc"
            | "hpp"
            | "html"
            | "htm"
            | "css"
            | "json"
            | "toml"
            | "yaml"
            | "yml"
            | "md"
            | "txt"
            | "sql"
            | "sh"
            | "ps1"
            | "lua"
            | "pas"
            | "pp"
            | "cob"
            | "cbl"
            | "f90"
            | "f95"
            | "r"
    )
}

pub(super) fn search_dir(
    root: &std::path::Path,
    needle: &str,
    hits: &mut Vec<super::ProjectSearchHit>,
    cap: usize,
) {
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    for entry in entries.flatten() {
        if hits.len() >= cap {
            return;
        }
        let path = entry.path();
        let name = path.file_name().unwrap_or_default().to_string_lossy();
        // Skip hidden entries and common build/cache dirs
        if name.starts_with('.') {
            continue;
        }
        if matches!(
            name.as_ref(),
            "target" | "node_modules" | "__pycache__" | "build" | "dist"
        ) {
            continue;
        }
        if path.is_dir() {
            search_dir(&path, needle, hits, cap);
        } else if path.is_file() {
            let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
            if !is_text_extension(ext) {
                continue;
            }
            let Ok(content) = std::fs::read_to_string(&path) else {
                continue;
            };
            let display = path.to_string_lossy().to_string();
            for (li, line) in content.lines().enumerate() {
                if hits.len() >= cap {
                    return;
                }
                if line.to_lowercase().contains(needle) {
                    hits.push(super::ProjectSearchHit {
                        file: display.clone(),
                        line: li,
                        snippet: line.trim().chars().take(120).collect(),
                    });
                }
            }
        }
    }
}
