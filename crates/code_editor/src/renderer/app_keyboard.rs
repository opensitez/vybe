use std::time::Instant;
use cosmic_text::{Attrs, Action, Cursor, Edit, Family, Motion, Selection, Shaping};
use winit::keyboard::{Key, NamedKey};
#[cfg(target_os = "macos")]
use winit::platform::modifier_supplement::KeyEventExtModifierSupplement;

use super::{App, TabContent};
use vybe_widgets::code_editor_widget::CodeEditorWidget;

impl App {
    pub(super) fn handle_key_press(&mut self, event: winit::event::KeyEvent) {
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

        // === Autocomplete keyboard intercepts ===
        if w.autocomplete_visible {
            match event.logical_key {
                Key::Named(NamedKey::ArrowUp) => {
                    if w.autocomplete_selected > 0 { w.autocomplete_selected -= 1; }
                    self.window.as_ref().unwrap().request_redraw();
                    return;
                }
                Key::Named(NamedKey::ArrowDown) => {
                    if w.autocomplete_selected < w.autocomplete_items.len().saturating_sub(1) { w.autocomplete_selected += 1; }
                    self.window.as_ref().unwrap().request_redraw();
                    return;
                }
                Key::Named(NamedKey::Enter) | Key::Named(NamedKey::Tab) => {
                    w.accept_autocomplete(&mut self.font_system);
                    w.sync();
                    self.pending_lsp_update = true;
                    self.last_lsp_update = Instant::now();
                    tab.is_modified = true;
                    self.window.as_ref().unwrap().request_redraw();
                    return;
                }
                Key::Named(NamedKey::Escape) => {
                    w.dismiss_autocomplete();
                    self.window.as_ref().unwrap().request_redraw();
                    return;
                }
                _ => {
                    w.dismiss_autocomplete();
                }
            }
        }

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
                    w.editor.action(&mut self.font_system, Action::Enter);
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
                     w.find_next(&mut self.font_system);
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
            Key::Named(NamedKey::F12) => {
                // Go to definition
                if let Some(path) = &tab.path {
                    let uri = format!("file://{}", path);
                    let line = w.editor.cursor().line as u32;
                    let col = w.editor.cursor().index as u32;
                    let _ = self.lsp.send(crate::lsp_client::LspRequest::Definition(uri, line, col));
                }
                self.window.as_ref().unwrap().request_redraw();
                return;
            }
            Key::Named(NamedKey::Space) if cmd => {
                // Ctrl+Space: trigger autocomplete
                self.trigger_completion();
                self.window.as_ref().unwrap().request_redraw();
                return;
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
                }

                for ch in t.chars() { if !ch.is_control() || ch == '\t' || ch == '\n' { if w.is_search_open { if w.is_replace_open && alt { w.replace_query.push(ch); } else { w.search_query.pop(); w.search_query.push(ch); } } else { 
                 let mut skip = false; if let Some(cl) = match ch { ')'=>Some(')'),'}'=>Some('}'),']'=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { 
                     let cli = w.editor.cursor().line; let cur = w.editor.cursor().index; 
                     let line_text = w.editor.with_buffer(|b| b.lines[cli].text().to_string());
                     let next_ch = line_text[cur..].chars().next();
                     if next_ch == Some(cl) { w.editor.action(&mut self.font_system, Action::Motion(Motion::Right)); skip = true; } 
                 }
                if !skip { w.editor.action(&mut self.font_system, Action::Insert(ch)); tab.is_modified = true; if let Some(cl) = match ch { '('=>Some(')'),'{'=>Some('}'),'['=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { w.editor.action(&mut self.font_system, Action::Insert(cl)); w.editor.action(&mut self.font_system, Action::Motion(Motion::Left)); } }
            } } }
                // Auto-trigger autocomplete after `.` or `::`
                let trigger = t.chars().last();
                if trigger == Some('.') { self.trigger_completion(); }
                else if trigger == Some(':') {
                    // Check if previous char was also `:` (i.e., `::`)
                    let w2 = match &self.tabs[self.active_tab].content { TabContent::Code(cw) => cw, _ => unreachable!() };
                    let is_double_colon = w2.editor.with_buffer(|b| {
                        let cli = w2.editor.cursor().line;
                        let cur = w2.editor.cursor().index;
                        let text = b.lines[cli].text();
                        cur >= 2 && text.get(cur-2..cur) == Some("::")
                    });
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
            self.window.as_ref().unwrap().request_redraw(); 
        }
    }
}
