use std::fs;
use std::path::Path;
use std::time::{Duration, Instant};
use cosmic_text::Edit;
use tiny_skia::Rect;
use winit::event::{ElementState, MouseButton, MouseScrollDelta};

use super::{App, Tab, TabContent, EditAction, SidebarTab, SCALE, TAB_BAR_HEIGHT, FOOTER_HEIGHT, SPLITTER_WIDTH, MINIMAP_WIDTH, SIDEBAR_TAB_H};
use crate::editor::Editor as MyEditor;
use crate::language::load_language;
use crate::lsp_client::LspRequest;
use vybe_widgets::{TreeEvent, Dropdown, DropdownEvent};
use vybe_widgets::code_editor_widget::CodeEditorWidget;

impl App {
    pub(super) fn handle_mouse_wheel(&mut self, delta: MouseScrollDelta) {
        let a = match delta {
            MouseScrollDelta::LineDelta(_, y) => y * 120.0,
            MouseScrollDelta::PixelDelta(pos) => pos.y as f32 * 2.0,
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

    pub(super) fn handle_cursor_moved(&mut self) {
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
                TabContent::Resources(_) => {}
            }
            needs_editor_redraw = true;
        }

        // Form designer menu open needs continuous redraw for hover
        let form_menu_open = self.tabs.iter().any(|t| matches!(&t.content, TabContent::Form(f) if f.menu_bar.open_menu.is_some()));

        // 6. LSP Hover Tooltip trigger (debounced)
        if !self.is_dragging && self.active_tab < self.tabs.len() {
            let ed_top_l = tch + TAB_BAR_HEIGHT;
            let ed_bottom_l = height - FOOTER_HEIGHT;
            if mx > ed_start_x && my > ed_top_l && my < ed_bottom_l {
                let dx = (mx - self.last_hover_pos.0).abs();
                let dy = (my - self.last_hover_pos.1).abs();
                // Clear tooltip if mouse moved
                if dx > 2.0 || dy > 2.0 {
                    if let TabContent::Code(cw) = &mut self.tabs[self.active_tab].content {
                        if cw.hover_text.is_some() {
                            cw.hover_text = None;
                            needs_editor_redraw = true;
                        }
                    }
                    self.last_hover_pos = (mx, my);
                    self.last_hover_time = Instant::now();
                }
                // Send hover request after 400ms dwell
                if self.last_hover_time.elapsed() >= Duration::from_millis(400) {
                    if let TabContent::Code(cw) = &self.tabs[self.active_tab].content {
                        if cw.hover_text.is_none() {
                            // Convert mouse position to buffer line/col
                            let rel_x = mx - ed_start_x;
                            let rel_y = my - ed_top_l + cw.scroll_y / SCALE;
                            let line_h = cw.editor.with_buffer(|b| b.metrics().line_height);
                            let line = (rel_y / line_h).max(0.0) as u32;
                            // Approximate column from x position (monospace ~9px per char at default size)
                            let gutter = 64.0; // GUTTER_WIDTH
                            let col = ((rel_x - gutter).max(0.0) / 9.0) as u32;
                            if let Some(path) = &self.tabs[self.active_tab].path {
                                let uri = format!("file://{}", path);
                                let _ = self.lsp.send(LspRequest::Hover(uri, line, col));
                            }
                            // Reset time so we don't spam requests
                            self.last_hover_time = Instant::now() - Duration::from_millis(200);
                        }
                    }
                }
            } else {
                // Mouse outside editor: clear hover
                if let TabContent::Code(cw) = &mut self.tabs[self.active_tab].content {
                    if cw.hover_text.is_some() {
                        cw.hover_text = None;
                        needs_editor_redraw = true;
                    }
                }
            }
        }

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

    pub(super) fn handle_mouse_input(&mut self, state: ElementState, button: MouseButton) {
        let mx = self.mouse_pos.0 / SCALE;
        let my = self.mouse_pos.1 / SCALE;
        let pw = self.pixmap.as_ref().unwrap().width() as f32;
        let ph = self.pixmap.as_ref().unwrap().height() as f32 / SCALE;
        let height = ph;

        if state == ElementState::Pressed && button == MouseButton::Right {
            // Right-click in project explorer sidebar → context menu
            let tch_r = self.top_chrome_h();
            if mx < self.explorer_width && my > tch_r && self.sidebar_tab == SidebarTab::Project {
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

            // 1. Minimap Hit-testing (code editor only)
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
                let menu_open = self.tabs.iter().any(|t| matches!(&t.content, TabContent::Form(ref f) if f.menu_bar.open_menu.is_some()));

                if my < tch || menu_open {
                    if let Some(form_tab) = self.tabs.iter_mut().find(|t| matches!(&t.content, TabContent::Form(_))) {
                        if let TabContent::Form(f) = &mut form_tab.content {
                            let menu_rect = crate::form_designer_tab::Rect { x: 0.0, y: 0.0, w: pw / SCALE, h: 28.0 };
                            if let Some(action) = f.menu_bar.handle_click(mx, my, menu_rect) {
                                self.window.as_ref().unwrap().request_redraw();
                                match action {
                                    crate::form_designer_tab::MenuAction::NewProject => {
                                        self.project = vybe_project::project::Project::new("Project1".to_string());
                                        let mut form = vybe_forms::Form::new("Form1".to_string());
                                        form.width = 640; form.height = 480;
                                        let fm = vybe_project::project::FormModule::new_classic(form);
                                        self.project.forms.push(fm);
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
                                        if self.project.resource_files.is_empty() {
                                            self.project.resource_files.push(vybe_project::ResourceManager::new());
                                        }
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
                                            let current_form_name = if self.active_tab < self.tabs.len() {
                                                let tab = &self.tabs[self.active_tab];
                                                if matches!(&tab.content, TabContent::Code(_)) {
                                                    tab.name.strip_suffix(".vb").map(|s| s.to_string())
                                                } else { None }
                                            } else { None };

                                            if let Some(ref fname) = current_form_name {
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
                                                if let Some(idx) = self.tabs.iter().position(|t| matches!(&t.content, TabContent::Form(_))) {
                                                    self.active_tab = idx;
                                                }
                                            }
                                        }
                                        crate::form_designer_tab::ToolbarAction::ViewCode => {
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

                                            if let Some(idx) = self.tabs.iter().position(|t| t.name == code_tab_name && matches!(&t.content, TabContent::Code(_))) {
                                                self.active_tab = idx;
                                            } else {
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

                        // Resources header
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
                        }
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
                            if r.handle_col_resize_start(mx, my, rx, ry, rw, rh) {
                                start_drag = true;
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
}
