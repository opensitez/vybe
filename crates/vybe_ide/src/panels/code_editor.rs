use egui::Ui;
use crate::state::EditorState;

fn highlight_vb(theme: &egui::Style, text: &str) -> egui::text::LayoutJob {
    let mut job = egui::text::LayoutJob::default();
    let font_id = egui::FontId::monospace(14.0);

    let is_dark = theme.visuals.dark_mode;
    let kw_color = if is_dark { egui::Color32::from_rgb(86, 156, 214) } else { egui::Color32::from_rgb(0, 0, 255) };
    let string_color = if is_dark { egui::Color32::from_rgb(206, 145, 120) } else { egui::Color32::from_rgb(163, 21, 21) };
    let comment_color = if is_dark { egui::Color32::from_rgb(106, 153, 85) } else { egui::Color32::from_rgb(0, 128, 0) };
    let default_color = theme.visuals.text_color();

    let keywords = [
        "Public", "Private", "Dim", "As", "Integer", "String", "Boolean", "Object",
        "If", "Then", "Else", "ElseIf", "End", "Function", "Sub", "Class", "Module",
        "Me", "True", "False", "New", "Handles", "ByVal", "ByRef", "For", "To", "Next",
        "While", "Return", "Property", "Get", "Set", "Select", "Case", "Try", "Catch"
    ];

    let mut in_comment = false;
    let mut in_string = false;
    let mut word = String::new();

    let mut flush_word = |job: &mut egui::text::LayoutJob, word: &mut String, color: egui::Color32| {
        if !word.is_empty() {
            job.append(word, 0.0, egui::text::TextFormat { font_id: font_id.clone(), color, ..Default::default() });
            word.clear();
        }
    };

    let mut chars = text.char_indices().peekable();
    while let Some((_, c)) = chars.next() {
        if in_comment {
            if c == '\n' || c == '\r' {
                flush_word(&mut job, &mut word, comment_color);
                in_comment = false;
                word.push(c);
                flush_word(&mut job, &mut word, default_color);
            } else {
                word.push(c);
            }
        } else if in_string {
            if c == '"' {
                word.push(c);
                flush_word(&mut job, &mut word, string_color);
                in_string = false;
            } else {
                word.push(c);
            }
        } else {
            if c == '\'' {
                if !word.is_empty() {
                    let is_kw = keywords.iter().any(|k| word.eq_ignore_ascii_case(k));
                    flush_word(&mut job, &mut word, if is_kw { kw_color } else { default_color });
                }
                in_comment = true;
                word.push(c);
            } else if c == '"' {
                if !word.is_empty() {
                    let is_kw = keywords.iter().any(|k| word.eq_ignore_ascii_case(k));
                    flush_word(&mut job, &mut word, if is_kw { kw_color } else { default_color });
                }
                in_string = true;
                word.push(c);
            } else if c.is_ascii_alphanumeric() || c == '_' {
                word.push(c);
            } else {
                if !word.is_empty() {
                    let is_kw = keywords.iter().any(|k| word.eq_ignore_ascii_case(k));
                    flush_word(&mut job, &mut word, if is_kw { kw_color } else { default_color });
                }
                word.push(c);
                flush_word(&mut job, &mut word, default_color);
            }
        }
    }

    if in_comment { flush_word(&mut job, &mut word, comment_color); }
    else if in_string { flush_word(&mut job, &mut word, string_color); }
    else if !word.is_empty() {
        let is_kw = keywords.iter().any(|k| word.eq_ignore_ascii_case(k));
        flush_word(&mut job, &mut word, if is_kw { kw_color } else { default_color });
    }

    job
}

pub fn show(ui: &mut Ui, state: &mut EditorState) {
    let name = state.current_code_file.clone()
        .or_else(|| state.current_form.clone());

    let Some(name) = name else {
        ui.label("No file selected.");
        return;
    };

    ui.horizontal(|ui| {
        ui.heading(format!("Code — {}", name));
    });
    ui.separator();

    let code = state.get_code_buffer(&name);
    
    // Tab to indent simple block
    let id = ui.id().with("code_edit");
    let mut consumed_tab = false;
    if ui.input(|i| i.key_pressed(egui::Key::Tab) && i.modifiers.is_none()) {
        if let Some(mut te_state) = egui::TextEdit::load_state(ui.ctx(), id) {
            if let Some(range) = te_state.cursor.char_range() {
                let primary = range.primary.index;
                let secondary = range.secondary.index;
                let min = primary.min(secondary);
                let max = primary.max(secondary);
                
                let selected_text: String = code.chars().skip(min).take(max - min).collect();
                if min != max && selected_text.contains('\n') {
                    // Find start of the first line involved
                    let mut start_of_line = min;
                    let chars: Vec<char> = code.chars().collect();
                    while start_of_line > 0 && chars[start_of_line - 1] != '\n' {
                        start_of_line -= 1;
                    }
                    
                    let target_block: String = chars[start_of_line..max].iter().collect();
                    let lines: Vec<&str> = target_block.split('\n').collect();
                    
                    let mut new_text = String::new();
                    let mut added_chars = 0;
                    for (i, line) in lines.iter().enumerate() {
                        if i > 0 { new_text.push('\n'); }
                        new_text.push_str("    ");
                        new_text.push_str(line);
                        added_chars += 4;
                    }
                    
                    code.replace_range(
                        code.char_indices().nth(start_of_line).map(|(i, _)| i).unwrap_or(0)..
                        code.char_indices().nth(max).map(|(i, _)| i).unwrap_or(code.len()),
                        &new_text
                    );
                    
                    let new_primary = primary + if primary == min { 4 } else { added_chars };
                    let new_secondary = secondary + if secondary == min { 4 } else { added_chars };
                    
                    te_state.cursor.set_char_range(Some(egui::text::CCursorRange::two(
                        egui::text::CCursor::new(new_primary),
                        egui::text::CCursor::new(new_secondary)
                    )));
                    te_state.store(ui.ctx(), id);
                    consumed_tab = true;
                }
            }
        }
    }
    
    if consumed_tab {
        ui.input_mut(|i| i.consume_key(egui::Modifiers::NONE, egui::Key::Tab));
    }

    egui::ScrollArea::both()
        .auto_shrink([false, false])
        .show(ui, |ui| {
            // Compute line numbers
            let num_lines = code.lines().count().max(1);
            let mut line_numbers = String::with_capacity(num_lines * 4);
            for i in 1..=num_lines {
                line_numbers.push_str(&format!("{:>3}\n", i));
            }

            ui.horizontal_top(|ui| {
                // Line numbers gutter
                let theme = ui.style().clone();
                let gutter_color = if theme.visuals.dark_mode { egui::Color32::from_rgb(133, 133, 133) } else { egui::Color32::from_rgb(150, 150, 150) };
                
                ui.add(
                    egui::TextEdit::multiline(&mut line_numbers.as_str())
                        .font(egui::TextStyle::Monospace)
                        .text_color(gutter_color)
                        .interactive(false)
                        .frame(false)
                        .desired_width(32.0)
                );
                ui.add_space(4.0);
                
                // Code editor
                let mut layouter = |ui: &egui::Ui, text: &str, wrap_width: f32| {
                    let mut layout_job = highlight_vb(&theme, text);
                    layout_job.wrap.max_width = wrap_width;
                    ui.fonts(|f| f.layout_job(layout_job))
                };

                let response = ui.add_sized(
                    ui.available_size(),
                    egui::TextEdit::multiline(code)
                        .id_source("code_edit")
                        .font(egui::TextStyle::Monospace)
                        .code_editor()
                        .lock_focus(true)
                        .frame(false)
                        .desired_width(f32::INFINITY)
                        .layouter(&mut layouter),
                );
            });
        });
}
