use std::time::Instant;
use winit::keyboard::{Key, NamedKey};

use super::{App, TabContent};
use vybe_widgets::CursorMotion;

impl App {
    pub(super) fn handle_key_press(&mut self, event: vybe_widgets::KeyEvent) {
        if self.tabs.is_empty() { return; }

        // ── Global overlays take precedence over any tab-specific handling ──
        if self.is_command_palette {
            self.command_palette_handle_key(&event);
            return;
        }
        if self.is_project_search {
            self.project_search_handle_key(&event);
            return;
        }

        // Handle Form tab keyboard events
        if let TabContent::Form(f) = &mut self.tabs[self.active_tab].content {
            // Inline property editing takes precedence so typing in the
            // value cell doesn't collide with form-level shortcuts.
            if f.handle_properties_key(&event) {
                self.sync_active_form_to_project();
                return;
            }
            let cmd = event.cmd;
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
                            self.project = vybex::projects::project::Project::new("Project1".to_string());
                            let mut form = vybex::projects::Form::new("Form1".to_string());
                            form.width = 640; form.height = 480;
                            self.project.forms.push(vybex::projects::project::FormModule::new_classic(form.clone()));
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
                                if let Ok(proj) = vybex::projects::serialization::load_project_auto(&path_str) {
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

            self.sync_active_form_to_project();
            return;
        }

        // Handle Resources tab keyboard events
        if let TabContent::Resources(r) = &mut self.tabs[self.active_tab].content {
            let key = &event.logical_key;
            let cmd = event.cmd;
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
            
            return;
        }

        let mut acted = true;
        let tab = &mut self.tabs[self.active_tab];
        let w = match &mut tab.content {
            TabContent::Code(cw) => cw,
            _ => {  return; }
        };
        let _theme = w.theme.clone();
        let cmd = event.cmd;
        let alt = event.alt; let shift = event.shift;

        // === Autocomplete keyboard intercepts ===
        if w.autocomplete_visible {
            match event.logical_key {
                Key::Named(NamedKey::ArrowUp) => {
                    if w.autocomplete_selected > 0 { w.autocomplete_selected -= 1; }
                    
                    return;
                }
                Key::Named(NamedKey::ArrowDown) => {
                    if w.autocomplete_selected < w.autocomplete_items.len().saturating_sub(1) { w.autocomplete_selected += 1; }
                    
                    return;
                }
                Key::Named(NamedKey::Enter) | Key::Named(NamedKey::Tab) => {
                    w.accept_autocomplete(&mut self.font_system);
                    w.sync();
                    self.pending_lsp_update = true;
                    self.last_lsp_update = Instant::now();
                    tab.is_modified = true;
                    
                    return;
                }
                Key::Named(NamedKey::Escape) => {
                    w.dismiss_autocomplete();
                    
                    return;
                }
                _ => {
                    w.dismiss_autocomplete();
                }
            }
        }

        let key_str = match event.key_without_modifiers.clone() {
            Key::Character(c) => c.to_lowercase(),
            Key::Named(nk) => format!("{:?}", nk),
            _ => String::new(),
        };

        for kb in &self.keybindings {
            if kb.key == key_str && kb.cmd == cmd && kb.shift == shift && kb.alt == alt {
                match kb.action.as_str() {
                    "Undo" => {
                        let (cl, ci) = w.cursor_pos();
                        if let Some((text, line, col)) = w.my_editor.undo(cl, ci) {
                            w.set_buffer_text(&mut self.font_system, &text);
                            w.set_cursor_pos(line, col);
                        }
                    }
                    "Redo" => {
                        let (cl, ci) = w.cursor_pos();
                        if let Some((text, line, col)) = w.my_editor.redo(cl, ci) {
                            w.set_buffer_text(&mut self.font_system, &text);
                            w.set_cursor_pos(line, col);
                        }
                    }
                    "SelectAll" => {
                        w.select_all();
                    }
                    "MoveBufferStart" => w.action_motion(&mut self.font_system, CursorMotion::BufferStart),
                    "MoveBufferEnd" => w.action_motion(&mut self.font_system, CursorMotion::BufferEnd),
                    "MoveLineStart" => w.action_motion(&mut self.font_system, CursorMotion::Home),
                    "MoveLineEnd" => w.action_motion(&mut self.font_system, CursorMotion::End),
                    "MoveWordLeft" => w.action_motion(&mut self.font_system, CursorMotion::LeftWord),
                    "MoveWordRight" => w.action_motion(&mut self.font_system, CursorMotion::RightWord),
                    "Save" => {
                        let save_path = tab.path.clone();
                        let text = w.my_editor.rope.to_string();
                        tab.is_modified = false;
                        if let Some(path) = save_path {
                            let msg = match std::fs::write(&path, &text) {
                                Ok(_) => format!("Saved: {}", path),
                                Err(e) => format!("Save error: {}", e),
                            };
                            self.output_lines_buffer.push(msg);
                            self.output_panel.set_output_lines(&self.output_lines_buffer);
                            self.output_panel.set_visible(true);
                        } else if let Some(pp) = self.project_path.clone() {
                            // Sync current tab code to project before save
                            let tab_name = tab.name.clone();
                            let code = w.my_editor.rope.to_string();
                            if let Some(form_name) = tab_name.strip_suffix(".vb") {
                                if let Some(fm) = self.project.forms.iter_mut().find(|fm| fm.form.name == form_name) {
                                    fm.set_user_code(code);
                                }
                            } else if let Some(cf) = self.project.code_files.iter_mut().find(|cf| cf.name == tab_name) {
                                cf.code = code;
                            }
                            let msg = match vybex::projects::serialization::save_project_auto(&self.project, &pp) {
                                Ok(_) => format!("Saved: {}", pp),
                                Err(e) => format!("Save error: {}", e),
                            };
                            self.output_lines_buffer.push(msg);
                            self.output_panel.set_output_lines(&self.output_lines_buffer);
                            self.output_panel.set_visible(true);
                        }
                    }
                    "Find" => { w.is_search_open = true; if let Some(t) = w.copy_selection_text() { if !t.is_empty() { w.search_query = t; } } else { w.search_query.clear(); } }
                    "Replace" => { w.is_search_open = true; w.is_replace_open = !w.is_replace_open; }
                    _ => { acted = false; }
                }
                if acted { w.needs_reshape = true; w.sync();  return; }
            }
        }

        match event.key_without_modifiers.clone() {
            Key::Character(c) if cmd && (c == "=" || c == "+") => { w.set_zoom(&mut self.font_system, 1.0); }
            Key::Character(c) if cmd && c == "-" => { w.set_zoom(&mut self.font_system, -1.0); }
            Key::Character(c) if cmd && c == "0" => { w.font_size = 14.0; w.set_zoom(&mut self.font_system, 0.0); }
            Key::Character(c) if cmd && (c == "w" || c == "W") => { w.show_whitespace = !w.show_whitespace; }
            Key::Character(c) if alt && (c == "z" || c == "Z") => { w.wrap_lines = !w.wrap_lines; w.needs_reshape = true; }
            Key::Character(c) if cmd && (c == "m" || c == "M") => { let (cl, ci) = w.cursor_pos(); if let Some(p) = w.my_editor.find_matching_bracket(cl, ci, &w.lang_def) { w.set_cursor_pos(p.0, p.1); } }
            Key::Character(c) if cmd && shift && (c == "p" || c == "P") => {
                self.is_command_palette = !self.is_command_palette;
                self.command_palette_query.clear();
                self.command_palette_selected = 0;
            }
            Key::Character(c) if cmd && !shift && (c == "p" || c == "P") => {
                self.is_quick_open = !self.is_quick_open;
                self.quick_open_query.clear();
            }
            Key::Character(c) if cmd && shift && (c == "f" || c == "F") => {
                self.is_project_search = !self.is_project_search;
                self.project_search_query.clear();
                self.project_search_results.clear();
                self.project_search_selected = 0;
            }
            Key::Character(c) if cmd && (c == "g" || c == "G") => { self.goto_line_open = !self.goto_line_open; self.goto_line_query.clear(); }
            Key::Character(c) if cmd && (c == "`") => { let v = self.output_panel.visible(); self.output_panel.set_visible(!v); }
            Key::Character(c) if cmd && shift && (c == "b" || c == "B") => {
                self.build_config = match self.build_config { super::BuildConfig::Debug => super::BuildConfig::Release, super::BuildConfig::Release => super::BuildConfig::Debug };
            }
            // Fold/Unfold: Cmd+Shift+[ folds current, Cmd+Shift+] unfolds current
            Key::Character(c) if cmd && shift && c == "[" => {
                let line = w.cursor_pos().0;
                if w.my_editor.folds.iter().any(|(s, _)| *s == line) && !w.my_editor.collapsed_starts.contains(&line) {
                    w.my_editor.toggle_fold(line);
                }
            }
            Key::Character(c) if cmd && shift && c == "]" => {
                let line = w.cursor_pos().0;
                if w.my_editor.collapsed_starts.contains(&line) {
                    w.my_editor.toggle_fold(line);
                }
            }
            // Fold All: Cmd+Shift+0
            Key::Character(c) if cmd && shift && c == "0" => {
                let fold_starts: Vec<usize> = w.my_editor.folds.iter().map(|(s, _)| *s).collect();
                for s in fold_starts { w.my_editor.collapsed_starts.insert(s); }
            }
            // Unfold All: Cmd+Shift+9
            Key::Character(c) if cmd && shift && c == "9" => {
                w.my_editor.collapsed_starts.clear();
            }
             Key::Named(NamedKey::Home) => {
                  let (cli, cur) = w.cursor_pos();
                  let line_text = w.line_text(cli);
                  let first_byte_idx = line_text.char_indices().find(|&(_, c)| !c.is_whitespace()).map(|(i, _)| i).unwrap_or(line_text.len());
                  if cur == first_byte_idx { w.action_motion(&mut self.font_system, CursorMotion::Home); }
                  else { w.set_cursor_pos(cli, first_byte_idx); }
             }
            Key::Named(NamedKey::End) => w.action_motion(&mut self.font_system, CursorMotion::End),
            Key::Character(c) if cmd && shift && (c == "k" || c == "K") => { w.action_motion(&mut self.font_system, CursorMotion::End); w.action_backspace(&mut self.font_system); w.action_motion(&mut self.font_system, CursorMotion::Home); let len = w.line_text(w.cursor_pos().0).len(); for _ in 0..len { w.action_delete(&mut self.font_system); } w.action_delete(&mut self.font_system); }
            Key::Named(NamedKey::Backspace) => if self.goto_line_open { self.goto_line_query.pop(); } else if w.is_search_open { if w.is_replace_open && alt { w.replace_query.pop(); } else { w.search_query.pop(); } } else { w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1); w.action_backspace(&mut self.font_system); tab.is_modified = true; }
            Key::Named(NamedKey::Delete) => { w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1); w.action_delete(&mut self.font_system); tab.is_modified = true; }
            Key::Named(NamedKey::Enter) => {
                if self.goto_line_open {
                    if let Ok(line_num) = self.goto_line_query.trim().parse::<usize>() {
                        let target = line_num.saturating_sub(1);
                        let max_line = w.line_count().saturating_sub(1);
                        let safe_line = target.min(max_line);
                        w.set_cursor_pos(safe_line, 0);
                        w.needs_reshape = true;
                    }
                    self.goto_line_open = false;
                    self.goto_line_query.clear();
                } else if self.is_quick_open {
                    self.is_quick_open = false;
                } else if w.is_search_open { w.find_next(&mut self.font_system); } 
                else { 
                    w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1);
                    let line_idx = w.cursor_pos().0;
                    let byte_off = w.compute_byte_offset(line_idx, w.cursor_pos().1);
                    w.my_editor.insert_newline(byte_off, &w.lang_def);
                    w.action_enter(&mut self.font_system);
                    w.needs_reshape = true; w.sync(); tab.is_modified = true;
                }
            }
            Key::Named(NamedKey::Escape) => { self.is_quick_open = false; self.goto_line_open = false; w.is_search_open = false; w.context_menu = None; }
            Key::Character(c) if cmd && (c == "c" || c == "C") => { if let Some(t) = w.copy_selection_text() { if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } } }
            Key::Character(c) if cmd && (c == "v" || c == "V") => { 
                 if let Some(cb) = &mut self.clipboard { 
                     if let Ok(t) = cb.get_text() { 
                         w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1);
                         
                         // Handle selection replacement
                         if let Some((start, end)) = w.selection_bounds() {
                             let s_off = w.compute_byte_offset(start.0, start.1);
                             let e_off = w.compute_byte_offset(end.0, end.1);
                             w.my_editor.delete_range(s_off.min(e_off), s_off.max(e_off), &w.lang_def);
                             w.action_delete(&mut self.font_system);
                         }

                         let byte_off = w.compute_byte_offset(w.cursor_pos().0, w.cursor_pos().1);
                         
                         let (new_line, new_col) = w.my_editor.insert_string(byte_off, &t, &w.lang_def);
                         
                         // Sync cosmic-text
                         w.set_buffer_text(&mut self.font_system, &w.my_editor.rope().to_string());
                         w.set_cursor_pos(new_line, new_col);
                         tab.is_modified = true; 
                     } 
                 } 
            }
            Key::Character(c) if cmd && (c == "x" || c == "X") => { if let Some(t) = w.copy_selection_text() { w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1); if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } w.action_delete(&mut self.font_system); tab.is_modified = true; } }
            Key::Character(c) if cmd && (c == "d" || c == "D") => {
                 w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1);
                 if !w.has_selection() {
                     let li = w.cursor_pos().0;
                     w.my_editor.duplicate_line(li);
                 } else if let Some(_t) = w.copy_selection_text() {
                     w.find_next(&mut self.font_system);
                 }
                 w.needs_reshape = true; tab.is_modified = true;
            }
            Key::Named(NamedKey::ArrowUp) if alt => { w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1); let li = w.cursor_pos().0; w.my_editor.move_line_up(li); w.needs_reshape = true; tab.is_modified = true; }
            Key::Named(NamedKey::ArrowDown) if alt => { w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1); let li = w.cursor_pos().0; if shift { w.my_editor.duplicate_line(li); } else { w.my_editor.move_line_down(li); } w.needs_reshape = true; tab.is_modified = true; }
            Key::Named(NamedKey::ArrowLeft) => {
                if shift { w.start_selection_if_none(); }
                w.action_motion(&mut self.font_system, CursorMotion::Left);
                if !shift { w.clear_selection(); }
            }
            Key::Named(NamedKey::ArrowRight) => {
                if shift { w.start_selection_if_none(); }
                w.action_motion(&mut self.font_system, CursorMotion::Right);
                if !shift { w.clear_selection(); }
            }
            Key::Named(NamedKey::ArrowUp) => {
                if shift { w.start_selection_if_none(); }
                w.action_motion(&mut self.font_system, CursorMotion::Up);
                if !shift { w.clear_selection(); }
            }
            Key::Named(NamedKey::ArrowDown) => {
                if shift { w.start_selection_if_none(); }
                w.action_motion(&mut self.font_system, CursorMotion::Down);
                if !shift { w.clear_selection(); }
            }
            Key::Character(c) if cmd && c == "z" => { 
                if shift {
                    if let Some((text, line, col)) = w.my_editor.redo(w.cursor_pos().0, w.cursor_pos().1) {
                        w.set_buffer_text(&mut self.font_system, &text);
                        let safe_line = line.min(w.line_count().saturating_sub(1));
                        let safe_col = if safe_line < w.line_count() { col.min(w.line_text(safe_line).len()) } else { 0 };
                        w.set_cursor_pos(safe_line, safe_col);
                    }
                } else {
                    if let Some((text, line, col)) = w.my_editor.undo(w.cursor_pos().0, w.cursor_pos().1) {
                        w.set_buffer_text(&mut self.font_system, &text);
                        let safe_line = line.min(w.line_count().saturating_sub(1));
                        let safe_col = if safe_line < w.line_count() { col.min(w.line_text(safe_line).len()) } else { 0 };
                        w.set_cursor_pos(safe_line, safe_col);
                    }
                }
            }
            Key::Character(c) if cmd && c == "a" => {
                w.select_all();
            }
            Key::Named(NamedKey::F12) => {
                // Go to definition
                if let Some(path) = &tab.path {
                    let uri = format!("file://{}", path);
                    let (line, col) = w.cursor_pos();
                    let _ = self.lsp.send(crate::lsp_client::LspRequest::Definition(uri, line as u32, col as u32));
                }
                
                return;
            }
            Key::Named(NamedKey::Space) if cmd => {
                // Ctrl+Space: trigger autocomplete
                self.trigger_completion();
                
                return;
            }
            _ => { if let Some(t) = event.text { if !cmd {
                // Route text input to open dialogs
                if self.goto_line_open {
                    for ch in t.chars() { if ch.is_ascii_digit() { self.goto_line_query.push(ch); } }
                    
                    return;
                }
                if self.is_quick_open {
                    for ch in t.chars() { if !ch.is_control() { self.quick_open_query.push(ch); } }
                    
                    return;
                }
                w.my_editor.save_snapshot(w.cursor_pos().0, w.cursor_pos().1);
                
                // Handling selection replacement on type
                if let Some((start, end)) = w.selection_bounds() {
                    let s_off = w.compute_byte_offset(start.0, start.1);
                    let e_off = w.compute_byte_offset(end.0, end.1);
                    w.my_editor.delete_range(s_off.min(e_off), s_off.max(e_off), &w.lang_def);
                }

                for ch in t.chars() { if !ch.is_control() || ch == '\t' || ch == '\n' { if w.is_search_open { if w.is_replace_open && alt { w.replace_query.push(ch); } else { w.search_query.pop(); w.search_query.push(ch); } } else { 
                 let mut skip = false; if let Some(cl) = match ch { ')'=>Some(')'),'}'=>Some('}'),']'=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { 
                     let (cli, cur) = w.cursor_pos(); 
                     let line_text = w.line_text(cli);
                     let next_ch = line_text[cur..].chars().next();
                     if next_ch == Some(cl) { w.action_motion(&mut self.font_system, CursorMotion::Right); skip = true; } 
                 }
                if !skip { w.action_insert(&mut self.font_system, ch); tab.is_modified = true; if let Some(cl) = match ch { '('=>Some(')'),'{'=>Some('}'),'['=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { w.action_insert(&mut self.font_system, cl); w.action_motion(&mut self.font_system, CursorMotion::Left); } }
            } } }
                // Auto-trigger autocomplete after `.` or `::`
                let trigger = t.chars().last();
                if trigger == Some('.') { self.trigger_completion(); }
                else if trigger == Some(':') {
                    // Check if previous char was also `:` (i.e., `::`)
                    let w2 = match &self.tabs[self.active_tab].content { TabContent::Code(cw) => cw, _ => unreachable!() };
                    let (cli, cur) = w2.cursor_pos();
                    let text = w2.line_text(cli);
                    let is_double_colon = cur >= 2 && text.get(cur-2..cur) == Some("::");
                    if is_double_colon { self.trigger_completion(); }
                }
            } else { acted = false; } } else { acted = false; } }
        }
        if acted { 
            if let TabContent::Code(w) = &mut self.tabs[self.active_tab].content {
                w.needs_reshape = true; 
                w.sync(); 
                w.hover_text = None;
                self.pending_lsp_update = true;
                self.last_lsp_update = Instant::now();
            }

        }
    }

    /// Command palette key handler. Fuzzy search + execute.
    fn command_palette_handle_key(&mut self, event: &vybe_widgets::KeyEvent) {
        if event.state != winit::event::ElementState::Pressed { return; }
        match &event.logical_key {
            Key::Named(NamedKey::Escape) => { self.is_command_palette = false; }
            Key::Named(NamedKey::Enter) => {
                let matches = self.command_palette_matches();
                let choice = matches.get(self.command_palette_selected).copied();
                self.is_command_palette = false;
                if let Some(idx) = choice {
                    let cmd = super::palette_commands()[idx].action;
                    self.execute_palette_command(cmd);
                }
            }
            Key::Named(NamedKey::ArrowDown) => {
                let len = self.command_palette_matches().len();
                if len > 0 {
                    self.command_palette_selected = (self.command_palette_selected + 1).min(len - 1);
                }
            }
            Key::Named(NamedKey::ArrowUp) => {
                self.command_palette_selected = self.command_palette_selected.saturating_sub(1);
            }
            Key::Named(NamedKey::Backspace) => {
                self.command_palette_query.pop();
                self.command_palette_selected = 0;
            }
            _ => {
                if let Some(t) = &event.text {
                    for ch in t.chars() {
                        if !ch.is_control() { self.command_palette_query.push(ch); }
                    }
                    self.command_palette_selected = 0;
                }
            }
        }
        self.needs_redraw = true;
    }

    /// Project-wide text-search key handler.
    fn project_search_handle_key(&mut self, event: &vybe_widgets::KeyEvent) {
        if event.state != winit::event::ElementState::Pressed { return; }
        match &event.logical_key {
            Key::Named(NamedKey::Escape) => { self.is_project_search = false; }
            Key::Named(NamedKey::Enter) => {
                // If results are already there and a row is selected, jump.
                if !self.project_search_results.is_empty() {
                    let hit = self.project_search_results[self.project_search_selected].clone();
                    self.is_project_search = false;
                    self.project_search_open_hit(&hit);
                } else {
                    // Run the search.
                    self.run_project_search();
                }
            }
            Key::Named(NamedKey::ArrowDown) => {
                let len = self.project_search_results.len();
                if len > 0 {
                    self.project_search_selected = (self.project_search_selected + 1).min(len - 1);
                }
            }
            Key::Named(NamedKey::ArrowUp) => {
                self.project_search_selected = self.project_search_selected.saturating_sub(1);
            }
            Key::Named(NamedKey::Backspace) => {
                self.project_search_query.pop();
                self.run_project_search();
            }
            _ => {
                if let Some(t) = &event.text {
                    for ch in t.chars() {
                        if !ch.is_control() { self.project_search_query.push(ch); }
                    }
                    self.run_project_search();
                }
            }
        }
        self.needs_redraw = true;
    }

    /// The fuzzy-matched, score-sorted indices into `palette_commands()`
    /// that are currently visible in the palette UI.
    pub(super) fn command_palette_matches(&self) -> Vec<usize> {
        use fuzzy_matcher::FuzzyMatcher;
        use fuzzy_matcher::skim::SkimMatcherV2;
        let matcher = SkimMatcherV2::default();
        let mut hits: Vec<(i64, usize)> = super::palette_commands()
            .iter()
            .enumerate()
            .filter_map(|(i, c)| {
                if self.command_palette_query.is_empty() {
                    Some((0, i))
                } else {
                    matcher.fuzzy_match(c.label, &self.command_palette_query).map(|s| (s, i))
                }
            })
            .collect();
        hits.sort_by(|a, b| b.0.cmp(&a.0));
        hits.into_iter().map(|(_, i)| i).collect()
    }
}
