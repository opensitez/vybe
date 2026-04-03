use cosmic_text::{Attrs, Action, Edit, Family, Cursor, Shaping};

use super::{App, Tab, TabContent, EditAction};
use crate::editor::Editor as MyEditor;
use crate::language::load_language;
use crate::lsp_client::LspRequest;
use vybe_widgets::code_editor_widget::CodeEditorWidget;

impl App {
    /// Push a line to the output panel and auto-show it.
    pub(crate) fn output_push(&mut self, line: String) {
        self.output_lines.push(line);
        self.output_visible = true;
        let total = self.output_lines.len() as f32 * 18.0;
        let visible = self.output_panel_height - 24.0;
        self.output_scroll_y = (total - visible).max(0.0);
    }

    pub(super) fn flush_code_to_project(&mut self) {
        for tab in &self.tabs {
            if let TabContent::Code(cw) = &tab.content {
                let code = cw.my_editor.rope.to_string();
                let tab_name = &tab.name;
                if let Some(form_name) = tab_name.strip_suffix(".vb") {
                    if let Some(fm) = self.project.forms.iter_mut().find(|fm| fm.form.name == form_name) {
                        fm.set_user_code(code);
                        continue;
                    }
                }
                if let Some(cf) = self.project.code_files.iter_mut().find(|cf| cf.name == *tab_name) {
                    cf.code = code;
                }
            }
            if let TabContent::Form(f) = &tab.content {
                if let Some(fm) = self.project.forms.iter_mut().find(|fm| fm.form.name == f.form.name) {
                    fm.form = f.form.clone();
                }
            }
        }
    }

    pub(super) fn save_project(&mut self) {
        self.flush_code_to_project();
        if let Some(path) = self.project_path.clone() {
            match vybe_project::serialization::save_project_auto(&self.project, &path) {
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
                match vybe_project::serialization::save_project_auto(&self.project, &path_str) {
                    Ok(_) => self.output_push(format!("Saved: {}", path_str)),
                    Err(e) => self.output_push(format!("Save error: {}", e)),
                }
            }
        }
    }

    pub(super) fn run_project(&mut self) {
        self.flush_code_to_project();
        let path = match self.project_path.clone() {
            Some(p) => p,
            None => { self.output_push("Save the project first.".to_string()); return; }
        };
        let _ = vybe_project::serialization::save_project_auto(&self.project, &path);
        let config_flag = match self.build_config {
            super::BuildConfig::Debug => "--debug",
            super::BuildConfig::Release => "--release",
        };
        let vybec = std::env::current_exe().ok()
            .and_then(|p| p.parent().map(|d| d.join("vybec")))
            .unwrap_or_else(|| std::path::PathBuf::from("vybec"));
        self.output_push(format!("Building ({})...", if self.build_config == super::BuildConfig::Debug { "Debug" } else { "Release" }));
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
                    for line in reader.lines().take(200) {
                        if let Ok(l) = line { self.output_push(l); }
                    }
                }
                if let Some(stderr) = child.stderr.take() {
                    use std::io::BufRead;
                    let reader = std::io::BufReader::new(stderr);
                    for line in reader.lines().take(200) {
                        if let Ok(l) = line { self.output_push(format!("ERR: {}", l)); }
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
        else { return };
        for path in paths {
            match vybe_project::load_form_vb(&path) {
                Ok(fm) => {
                    let name = fm.form.name.clone();
                    if self.project.forms.iter().all(|f| f.form.name != name) {
                        let form_clone = fm.form.clone();
                        self.project.forms.push(fm);
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

    pub(super) fn add_existing_code(&mut self) {
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
                let lang = load_language("vb").or_else(|| load_language("rust")).expect("language not found");
                let my_editor = MyEditor::from_text(&code, &lang);
                let uri = format!("file:///project/{}", name);
                let widget = {
                    let text = my_editor.rope.to_string();
                    self.lsp.send(LspRequest::Init(text, "vb".to_string(), uri));
                    CodeEditorWidget::new(my_editor.inner, &mut self.font_system)
                };
                self.tabs.push(Tab { name, path: None, content: TabContent::Code(widget), is_sticky: true, buffer: None, is_modified: false });
                self.active_tab = self.tabs.len() - 1;
            }
        }
    }

    pub(super) fn remove_project_item(&mut self, name: &str) {
        let removed = self.project.remove_form(name) || self.project.remove_code_file(name);
        if removed {
            let tab_name_vb = format!("{}.vb", name);
            self.tabs.retain(|t| t.name != name && t.name != tab_name_vb);
            if self.active_tab >= self.tabs.len() && !self.tabs.is_empty() {
                self.active_tab = self.tabs.len() - 1;
            }
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

    pub(super) fn dispatch_edit_action(&mut self, action: EditAction) {
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
                    _ => {}
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

    pub(super) fn create_resource_editor_from_project(project: &vybe_project::project::Project) -> vybe_widgets::ResourceEditor {
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

    pub(super) fn process_resource_event(evt: vybe_widgets::ResourceEditorEvent, r: &mut vybe_widgets::ResourceEditor, project: &mut vybe_project::project::Project) {
        match evt {
            vybe_widgets::ResourceEditorEvent::AddResource(tab) => {
                if tab.is_file_based() {
                    let mut dialog = rfd::FileDialog::new();
                    let (filter_name, exts): (&str, Vec<&str>) = match tab {
                        vybe_widgets::ResourceTab::Images => ("Images", vec!["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"]),
                        vybe_widgets::ResourceTab::Icons => ("Icons", vec!["ico"]),
                        vybe_widgets::ResourceTab::Audio => ("Audio", vec!["wav", "mp3", "ogg", "flac", "aiff"]),
                        _ => ("All Files", vec!["*"]),
                    };
                    if !exts.is_empty() && exts[0] != "*" { dialog = dialog.add_filter(filter_name, &exts); }
                    dialog = dialog.add_filter("All Files", &["*"]);
                    if let Some(paths) = dialog.pick_files() {
                        for p in paths {
                            let path_str = p.to_string_lossy().to_string();
                            let name = p.file_stem().and_then(|s| s.to_str()).unwrap_or("Resource1")
                                .replace(|c: char| !c.is_alphanumeric() && c != '_', "_");
                            let res_tab = tab;
                            r.entries.push(vybe_widgets::ResourceEntry { name: name.clone(), value: path_str.clone(), comment: String::new(), tab: res_tab, file_name: Some(path_str.clone()) });
                            let rt = match res_tab {
                                vybe_widgets::ResourceTab::Images => vybe_project::resources::ResourceType::Image,
                                vybe_widgets::ResourceTab::Icons => vybe_project::resources::ResourceType::Icon,
                                vybe_widgets::ResourceTab::Audio => vybe_project::resources::ResourceType::Audio,
                                vybe_widgets::ResourceTab::Files => vybe_project::resources::ResourceType::File,
                                _ => vybe_project::resources::ResourceType::String,
                            };
                            if project.resource_files.is_empty() { project.resource_files.push(vybe_project::ResourceManager::new()); }
                            if let Some(rm) = project.resource_files.first_mut() { rm.resources.push(vybe_project::ResourceItem::new_file(name, &path_str, rt)); }
                        }
                    }
                } else {
                    r.entries.push(vybe_widgets::ResourceEntry { name: format!("NewResource{}", r.entries.len() + 1), value: String::new(), comment: String::new(), tab, file_name: None });
                }
                r.dirty = true;
            }
            vybe_widgets::ResourceEditorEvent::DeleteResource(idx) => {
                if idx < r.entries.len() { r.entries.remove(idx); r.selected_row = None; r.dirty = true; Self::sync_resources_to_project(r, project); }
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
                    if !exts.is_empty() { dialog = dialog.add_filter("Supported", &exts); }
                    dialog = dialog.add_filter("All Files", &["*"]);
                    if let Some(path) = dialog.pick_file() {
                        let path_str = path.to_string_lossy().to_string();
                        let name = path.file_stem().and_then(|s| s.to_str()).unwrap_or("Resource1").replace(|c: char| !c.is_alphanumeric() && c != '_', "_");
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
                r.entries.push(vybe_widgets::ResourceEntry { name: name.clone(), value: value.clone(), comment: comment.clone(), tab, file_name: None });
                r.dirty = true;
                if project.resource_files.is_empty() { project.resource_files.push(vybe_project::ResourceManager::new()); }
                if let Some(rm) = project.resource_files.first_mut() {
                    let mut item = vybe_project::ResourceItem::new_string(name, value);
                    item.resource_type = match tab { vybe_widgets::ResourceTab::Other => vybe_project::resources::ResourceType::Other, _ => vybe_project::resources::ResourceType::String };
                    item.comment = if comment.is_empty() { None } else { Some(comment) };
                    rm.resources.push(item);
                }
            }
            vybe_widgets::ResourceEditorEvent::EditCommitted(_, _, _) => { Self::sync_resources_to_project(r, project); }
            _ => {}
        }
    }

    pub(super) fn sync_resources_to_project(r: &vybe_widgets::ResourceEditor, project: &mut vybe_project::project::Project) {
        if project.resource_files.is_empty() { project.resource_files.push(vybe_project::ResourceManager::new()); }
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

    /// Send a completion request to the LSP for the current cursor position.
    pub(super) fn trigger_completion(&self) {
        if let Some(tab) = self.tabs.get(self.active_tab) {
            if let TabContent::Code(w) = &tab.content {
                let cursor = w.editor.cursor();
                let uri = tab.path.as_deref()
                    .map(|p| format!("file://{}", p))
                    .unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", tab.name));
                self.lsp.send(LspRequest::Completion(uri, cursor.line as u32, cursor.index as u32));
            }
        }
    }
}
