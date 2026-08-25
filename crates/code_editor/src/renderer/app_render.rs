use fuzzy_matcher::FuzzyMatcher;
use fuzzy_matcher::skim::SkimMatcherV2;
use tiny_skia::{Paint, PathBuilder, Pixmap, Rect, Stroke, Transform};

use super::{
    App, EditAction, FOOTER_HEIGHT, GUTTER_WIDTH, Keybinding, MINIMAP_WIDTH, PaletteAction, SCALE,
    SIDEBAR_TAB_H, SPLITTER_WIDTH, SidebarTab, TAB_BAR_HEIGHT, TabContent, UI_BAR_HEIGHT,
};
use crate::form_designer_tab::MenuAction;
use crate::lsp_client::{LspEvent, LspRequest};
use widgets::PanelWidget;
use widgets::TextColor;
use widgets::layout::RenderContext;
use widgets::output_panel::{ProblemEntry, ProblemSeverity};

fn format_kb(kb: &Keybinding) -> String {
    let mut s = String::new();
    if kb.cmd {
        s.push('⌘');
    }
    if kb.shift {
        s.push('⇧');
    }
    if kb.alt {
        s.push('⌥');
    }
    let key = match kb.key.as_str() {
        "ArrowUp" => "↑",
        "ArrowDown" => "↓",
        "ArrowLeft" => "←",
        "ArrowRight" => "→",
        "Enter" => "↵",
        "Tab" => "⇥",
        "Backspace" => "⌫",
        "Delete" => "⌦",
        "Escape" => "⎋",
        "Space" => "␣",
        k => k,
    };
    s.push_str(key);
    s
}

fn kb_hint_for_palette(kbs: &[Keybinding], action: PaletteAction) -> String {
    let target: Option<&str> = match action {
        PaletteAction::Edit(EditAction::Undo) => Some("Undo"),
        PaletteAction::Edit(EditAction::Redo) => Some("Redo"),
        PaletteAction::Edit(EditAction::Cut) => Some("Cut"),
        PaletteAction::Edit(EditAction::Copy) => Some("Copy"),
        PaletteAction::Edit(EditAction::Paste) => Some("Paste"),
        PaletteAction::Edit(EditAction::Delete) => None,
        PaletteAction::Menu(MenuAction::SaveProject) | PaletteAction::Menu(MenuAction::SaveAs) => {
            Some("Save")
        }
        PaletteAction::FindInFile => Some("Find"),
        PaletteAction::FindInProject => None,
        PaletteAction::GoToLine => None,
        PaletteAction::ToggleOutput => Some("ToggleOutput"),
        PaletteAction::CloseTab => None,
        PaletteAction::NextTab => None,
        PaletteAction::PrevTab => None,
        _ => None,
    };
    if let Some(action_str) = target {
        if let Some(kb) = kbs.iter().find(|k| k.action == action_str) {
            return format_kb(kb);
        }
    }
    String::new()
}

impl App {
    pub(super) fn render_internal(&mut self, pix: &mut Pixmap) {
        self.process_widget_events();

        // Debounce LSP Update
        if self.pending_lsp_update && self.last_lsp_update.elapsed().as_millis() > 300 {
            let mut lsp_text = None;
            if let Some(tab) = self.tabs.get(self.active_tab) {
                if let TabContent::Code(cw) = &tab.content {
                    let text = cw.my_editor.rope.to_string();
                    let uri = App::tab_uri(tab);
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
        let pix = pix;

        let to_skia = |c: TextColor| tiny_skia::Color::from_rgba8(c.r(), c.g(), c.b(), c.a());
        pix.fill(to_skia(theme.bg));

        // Compute top chrome height (menu + toolbar when any Form tab exists)
        let has_form_tab = self
            .tabs
            .iter()
            .any(|t| matches!(&t.content, TabContent::Form(_)));
        let top_chrome_h: f32 = if has_form_tab { 28.0 + 36.0 } else { 0.0 };
        let top_chrome_px = top_chrome_h * SCALE;

        // 0. Menu bar + Toolbar (always present when a Form tab exists)
        if has_form_tab {
            if let Some(form_tab) = self
                .tabs
                .iter()
                .find(|t| matches!(&t.content, TabContent::Form(_)))
            {
                if let TabContent::Form(f) = &form_tab.content {
                    let menu_rect = crate::form_designer_tab::Rect {
                        x: 0.0,
                        y: 0.0,
                        w: pix.width() as f32 / SCALE,
                        h: 28.0,
                    };
                    let tb_rect = crate::form_designer_tab::Rect {
                        x: 0.0,
                        y: 28.0,
                        w: pix.width() as f32 / SCALE,
                        h: 36.0,
                    };
                    f.menu_bar.render(
                        pix,
                        &mut self.font_system,
                        &mut self.swash_cache,
                        menu_rect,
                        SCALE,
                    );
                    crate::form_designer_tab::render_toolbar_pub(
                        pix,
                        &mut self.font_system,
                        &mut self.swash_cache,
                        tb_rect,
                        SCALE,
                    );
                }
            }
        }

        // 1. Sidebar
        let sidebar_x = 0.0;
        let sidebar_top = top_chrome_px;
        let sidebar_w = self.explorer_width * SCALE;
        let sidebar_h = pix.height() as f32 - top_chrome_px;

        // Sidebar background
        let mut sp = Paint::default();
        sp.set_color_rgba8(
            theme.sidebar_bg.r(),
            theme.sidebar_bg.g(),
            theme.sidebar_bg.b(),
            theme.sidebar_bg.a(),
        );
        pix.fill_rect(
            Rect::from_xywh(sidebar_x, sidebar_top, sidebar_w, sidebar_h).unwrap(),
            &sp,
            Transform::identity(),
            None,
        );

        // Sidebar tabs (Files | Project) — rendered by toolkit TabPanel
        let stab_h = SIDEBAR_TAB_H * SCALE;
        {
            // Sync sidebar tabs active state
            let sidebar_active = match self.sidebar_tab {
                SidebarTab::Files => 0,
                SidebarTab::Project => 1,
            };
            self.sidebar_tabs.set_active(sidebar_active);
            self.sidebar_tabs.set_colors(
                (
                    theme.sidebar_bg.r(),
                    theme.sidebar_bg.g(),
                    theme.sidebar_bg.b(),
                    theme.sidebar_bg.a(),
                ),
                (
                    theme.inactive_tab_bg.r(),
                    theme.inactive_tab_bg.g(),
                    theme.inactive_tab_bg.b(),
                    theme.inactive_tab_bg.a(),
                ),
                (
                    theme.active_tab_bg.r(),
                    theme.active_tab_bg.g(),
                    theme.active_tab_bg.b(),
                    theme.active_tab_bg.a(),
                ),
                (
                    theme.inactive_tab_text.r(),
                    theme.inactive_tab_text.g(),
                    theme.inactive_tab_text.b(),
                    theme.inactive_tab_text.a(),
                ),
                (
                    theme.active_tab_text.r(),
                    theme.active_tab_text.g(),
                    theme.active_tab_text.b(),
                    theme.active_tab_text.a(),
                ),
                (theme.kw.r(), theme.kw.g(), theme.kw.b(), 255),
            );
            let mut ctx = RenderContext {
                pixmap: pix,
                font_system: &mut self.font_system,
                swash_cache: &mut self.swash_cache,
                scale: SCALE,
            };
            self.sidebar_tabs.render(&mut ctx);
        }
        // Tab separator
        {
            let mut lp = Paint::default();
            lp.set_color_rgba8(
                theme.splitter_bg.r(),
                theme.splitter_bg.g(),
                theme.splitter_bg.b(),
                255,
            );
            pix.fill_rect(
                Rect::from_xywh(sidebar_x, sidebar_top + stab_h - 1.0, sidebar_w, 1.0).unwrap(),
                &lp,
                Transform::identity(),
                None,
            );
        }

        // Sidebar content below tabs
        let sidebar_content_y = sidebar_top + stab_h;
        match self.sidebar_tab {
            SidebarTab::Files => {
                if let Some(tab) = self.tabs.get(self.active_tab) {
                    if let Some(path) = &tab.path {
                        self.tree_view.reveal_path(path);
                    }
                }
                self.tree_view.render_tree(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    sidebar_x,
                    sidebar_content_y,
                    sidebar_w,
                    theme.sidebar_text,
                    (
                        theme.selection.r(),
                        theme.selection.g(),
                        theme.selection.b(),
                        theme.selection.a(),
                    ),
                );
            }
            SidebarTab::Project => {
                self.render_project_explorer(
                    pix,
                    sidebar_x,
                    sidebar_content_y,
                    sidebar_w,
                    sidebar_h - stab_h,
                    &theme,
                );
            }
        }

        // 1b. Splitter
        let mut slp = Paint::default();
        if self.is_dragging_splitter {
            slp.set_color_rgba8(0, 122, 204, 255);
        } else if self.hovering_splitter {
            slp.set_color_rgba8(
                theme.splitter_bg.r(),
                theme.splitter_bg.g(),
                theme.splitter_bg.b(),
                255,
            );
        } else {
            slp.set_color_rgba8(
                theme.splitter_bg.r(),
                theme.splitter_bg.g(),
                theme.splitter_bg.b(),
                255,
            );
        }
        pix.fill_rect(
            Rect::from_xywh(
                self.explorer_width * SCALE,
                top_chrome_px,
                SPLITTER_WIDTH * SCALE,
                pix.height() as f32 - top_chrome_px,
            )
            .unwrap(),
            &slp,
            Transform::identity(),
            None,
        );

        let mut lp = Paint::default();
        lp.set_color_rgba8(
            theme.splitter_bg.r().saturating_add(20),
            theme.splitter_bg.g().saturating_add(20),
            theme.splitter_bg.b().saturating_add(20),
            255,
        );
        pix.fill_rect(
            Rect::from_xywh(
                (self.explorer_width + SPLITTER_WIDTH) * SCALE,
                top_chrome_px,
                1.0 * SCALE,
                pix.height() as f32 - top_chrome_px,
            )
            .unwrap(),
            &lp,
            Transform::identity(),
            None,
        );

        // 2. Tab Bar — rendered by toolkit TabPanel
        let ed_start_x = (self.explorer_width + SPLITTER_WIDTH + 1.0) * SCALE;
        {
            // Update tab panel theme colors
            self.tab_panel.set_colors(
                (theme.bg.r(), theme.bg.g(), theme.bg.b(), theme.bg.a()),
                (
                    theme.tab_bar_bg.r(),
                    theme.tab_bar_bg.g(),
                    theme.tab_bar_bg.b(),
                    theme.tab_bar_bg.a(),
                ),
                (
                    theme.active_tab_bg.r(),
                    theme.active_tab_bg.g(),
                    theme.active_tab_bg.b(),
                    theme.active_tab_bg.a(),
                ),
                (
                    theme.inactive_tab_text.r(),
                    theme.inactive_tab_text.g(),
                    theme.inactive_tab_text.b(),
                    theme.inactive_tab_text.a(),
                ),
                (
                    theme.active_tab_text.r(),
                    theme.active_tab_text.g(),
                    theme.active_tab_text.b(),
                    theme.active_tab_text.a(),
                ),
                (theme.kw.r(), theme.kw.g(), theme.kw.b(), 255),
            );
            let mut ctx = RenderContext {
                pixmap: pix,
                font_system: &mut self.font_system,
                swash_cache: &mut self.swash_cache,
                scale: SCALE,
            };
            self.tab_panel.render(&mut ctx);
        }

        // 2b. Breadcrumb bar — file path above the editor
        {
            let br_top = top_chrome_px + TAB_BAR_HEIGHT * SCALE;
            let br_x = ed_start_x;
            let br_w = pix.width() as f32 - ed_start_x;
            let br_h = UI_BAR_HEIGHT * SCALE;
            // Background
            let mut bg = Paint::default();
            bg.set_color_rgba8(
                theme.footer_bg.r(),
                theme.footer_bg.g(),
                theme.footer_bg.b(),
                theme.footer_bg.a(),
            );
            pix.fill_rect(
                Rect::from_xywh(br_x, br_top, br_w, br_h).unwrap(),
                &bg,
                Transform::identity(),
                None,
            );
            // Bottom divider
            let mut dp = Paint::default();
            dp.set_color_rgba8(
                theme.guide.r(),
                theme.guide.g(),
                theme.guide.b(),
                theme.guide.a(),
            );
            pix.fill_rect(
                Rect::from_xywh(br_x, br_top + br_h - SCALE, br_w, SCALE).unwrap(),
                &dp,
                Transform::identity(),
                None,
            );
            // Crumbs text
            let crumbs = self
                .tabs
                .get(self.active_tab)
                .map(|t| {
                    let proj = &self.project.name;
                    let tname = &t.name;
                    format!("{}  \u{203A}  {}", proj, tname)
                })
                .unwrap_or_default();
            App::draw_ui_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &crumbs,
                br_x + 10.0 * SCALE,
                br_top + 4.0 * SCALE,
                theme.inactive_tab_text,
            );
        }

        // 3. Active Editor or Designer
        let output_h = if self.output_panel.visible() {
            self.output_panel_height
        } else {
            0.0
        };
        if self.active_tab < self.tabs.len() {
            let ed_top = top_chrome_px + (TAB_BAR_HEIGHT + UI_BAR_HEIGHT) * SCALE;
            let rect = Rect::from_xywh(
                ed_start_x,
                ed_top,
                pix.width() as f32 - ed_start_x,
                pix.height() as f32 - (ed_top + (FOOTER_HEIGHT + output_h) * SCALE),
            )
            .unwrap();

            // Drain LSP events
            while let Ok(evt) = self.lsp.rx.try_recv() {
                match evt {
                    LspEvent::Diagnostics(uri, diags) => {
                        for t in &mut self.tabs {
                            let t_uri = App::tab_uri(t);
                            if t_uri == uri {
                                if let TabContent::Code(cw) = &mut t.content {
                                    cw.my_editor.diagnostics = diags
                                        .iter()
                                        .map(|d| widgets::DiagnosticInfo {
                                            line: d.range.start.line as usize,
                                            col_start: d.range.start.character as usize,
                                            col_end: d.range.end.character as usize,
                                            message: d.message.clone(),
                                            severity: match d.severity {
                                                Some(lsp_types::DiagnosticSeverity::ERROR) => {
                                                    widgets::DiagnosticSeverity::Error
                                                }
                                                Some(lsp_types::DiagnosticSeverity::WARNING) => {
                                                    widgets::DiagnosticSeverity::Warning
                                                }
                                                Some(
                                                    lsp_types::DiagnosticSeverity::INFORMATION,
                                                ) => widgets::DiagnosticSeverity::Info,
                                                _ => widgets::DiagnosticSeverity::Hint,
                                            },
                                        })
                                        .collect();
                                    self.needs_redraw = true;
                                    break;
                                }
                            }
                        }
                    }
                    LspEvent::Completion(items) => {
                        use lsp_types::CompletionItemKind;
                        use widgets::code_editor_widget::AutocompleteItem;
                        if let Some(tab) = self.tabs.get_mut(self.active_tab) {
                            if let TabContent::Code(cw) = &mut tab.content {
                                cw.autocomplete_items = items
                                    .into_iter()
                                    .map(|ci| {
                                        let kind_icon = match ci.kind {
                                            Some(CompletionItemKind::FUNCTION)
                                            | Some(CompletionItemKind::METHOD) => "fn",
                                            Some(CompletionItemKind::STRUCT)
                                            | Some(CompletionItemKind::CLASS)
                                            | Some(CompletionItemKind::INTERFACE) => "ty",
                                            Some(CompletionItemKind::KEYWORD) => "kw",
                                            Some(CompletionItemKind::MODULE) => "md",
                                            Some(CompletionItemKind::FIELD)
                                            | Some(CompletionItemKind::PROPERTY) => "fi",
                                            Some(CompletionItemKind::VARIABLE) => "va",
                                            Some(CompletionItemKind::CONSTANT)
                                            | Some(CompletionItemKind::ENUM_MEMBER) => "ct",
                                            Some(CompletionItemKind::SNIPPET) => "sn",
                                            Some(CompletionItemKind::ENUM) => "en",
                                            Some(CompletionItemKind::TYPE_PARAMETER) => "tp",
                                            _ => "  ",
                                        };
                                        AutocompleteItem {
                                            label: ci.label,
                                            detail: ci.detail,
                                            insert_text: ci.insert_text,
                                            kind_icon,
                                        }
                                    })
                                    .collect();
                                cw.autocomplete_selected = 0;
                                cw.autocomplete_visible = !cw.autocomplete_items.is_empty();
                                self.needs_redraw = true;
                            }
                        }
                    }
                    LspEvent::Hover(_uri, text) => {
                        if let Some(tab) = self.tabs.get_mut(self.active_tab) {
                            if let TabContent::Code(cw) = &mut tab.content {
                                cw.hover_text = if text.is_empty() { None } else { Some(text) };
                                cw.hover_pos = self.mouse_pos;
                                self.needs_redraw = true;
                            }
                        }
                    }
                    LspEvent::Definition(uri, pos) => {
                        // Navigate to definition: find or open the file, jump to position
                        let target_path = uri.strip_prefix("file://").unwrap_or(&uri);
                        let target_name = std::path::Path::new(target_path)
                            .file_name()
                            .map(|n| n.to_string_lossy().to_string())
                            .unwrap_or_else(|| target_path.to_string());
                        // Check if tab already open
                        let mut found = None;
                        for (i, t) in self.tabs.iter().enumerate() {
                            if t.name == target_name || t.path.as_deref() == Some(target_path) {
                                found = Some(i);
                                break;
                            }
                        }
                        if let Some(idx) = found {
                            self.active_tab = idx;
                            if let TabContent::Code(cw) = &mut self.tabs[idx].content {
                                cw.set_cursor_pos(pos.line as usize, pos.character as usize);
                                cw.needs_reshape = true;
                            }
                        }
                        self.needs_redraw = true;
                    }
                }
            }

            let tab = &mut self.tabs[self.active_tab];
            match &mut tab.content {
                TabContent::Form(f) => {
                    f.render(
                        pix,
                        &mut self.font_system,
                        &mut self.swash_cache,
                        crate::form_designer_tab::Rect {
                            x: rect.left() / SCALE,
                            y: rect.top() / SCALE,
                            w: rect.width() / SCALE,
                            h: rect.height() / SCALE,
                        },
                        SCALE,
                    );
                }
                TabContent::Code(w) => {
                    w.sync_wrap_and_size(
                        &mut self.font_system,
                        rect.width() - (GUTTER_WIDTH + MINIMAP_WIDTH) * SCALE,
                        rect.height(),
                    );
                    w.render_pixels(pix, &mut self.font_system, &mut self.swash_cache, rect);
                }
                TabContent::Resources(r) => {
                    r.render_at(
                        pix,
                        &mut self.font_system,
                        &mut self.swash_cache,
                        rect.left() / SCALE,
                        rect.top() / SCALE,
                        rect.width() / SCALE,
                        rect.height() / SCALE,
                        SCALE,
                    );
                }
            }
        }

        // 3b. Output / Problems Panel — rendered by toolkit OutputPanel
        {
            // Sync problems from all tabs into the OutputPanel
            let problems: Vec<ProblemEntry> = self
                .tabs
                .iter()
                .flat_map(|t| {
                    let name = t.name.clone();
                    if let TabContent::Code(cw) = &t.content {
                        cw.my_editor
                            .diagnostics
                            .iter()
                            .map(move |d| ProblemEntry {
                                file: name.clone(),
                                line: d.line + 1,
                                severity: match d.severity {
                                    widgets::DiagnosticSeverity::Error => {
                                        ProblemSeverity::Error
                                    }
                                    widgets::DiagnosticSeverity::Warning => {
                                        ProblemSeverity::Warning
                                    }
                                    widgets::DiagnosticSeverity::Info => ProblemSeverity::Info,
                                    widgets::DiagnosticSeverity::Hint => ProblemSeverity::Hint,
                                },
                                message: d.message.clone(),
                            })
                            .collect::<Vec<_>>()
                    } else {
                        vec![]
                    }
                })
                .collect();
            self.output_panel.set_problems(problems);

            let mut ctx = RenderContext {
                pixmap: pix,
                font_system: &mut self.font_system,
                swash_cache: &mut self.swash_cache,
                scale: SCALE,
            };
            self.output_panel.render(&mut ctx);
        }

        // 4. Footer — rendered by toolkit StatusBarPanel
        {
            // Update status bar sections based on current state
            self.status_bar.sections_mut().clear();
            self.status_bar.set_background(
                theme.footer_bg.r(),
                theme.footer_bg.g(),
                theme.footer_bg.b(),
                theme.footer_bg.a(),
            );

            // Left section: diagnostic summary — clickable to open Problems tab.
            {
                let mut errors = 0usize;
                let mut warnings = 0usize;
                for t in &self.tabs {
                    if let TabContent::Code(cw) = &t.content {
                        for d in &cw.my_editor.diagnostics {
                            match d.severity {
                                widgets::DiagnosticSeverity::Error => errors += 1,
                                widgets::DiagnosticSeverity::Warning => warnings += 1,
                                _ => {}
                            }
                        }
                    }
                }
                let diag_label = format!("\u{2715} {}   \u{26A0} {}", errors, warnings);
                self.status_bar.add_section_with_id(
                    &diag_label,
                    diag_label.len() as f32 * 8.5,
                    false,
                    "diagnostics",
                );
            }

            // Left section: cursor info
            let status_text = if let Some(tab) = self.tabs.get(self.active_tab) {
                match &tab.content {
                    TabContent::Code(cw) => {
                        let cursor = cw.cursor_pos();
                        let text = cw.my_editor.rope.to_string();
                        let line_endings = if text.contains("\r\n") { "CRLF" } else { "LF" };
                        let zoom_pct = (cw.font_size / 14.0 * 100.0) as i32;
                        let ws = if cw.show_whitespace { " | WS" } else { "" };
                        format!(
                            "Ln {}, Col {} | {}% | {}{} | UTF-8",
                            cursor.0 + 1,
                            cursor.1 + 1,
                            zoom_pct,
                            line_endings,
                            ws
                        )
                    }
                    TabContent::Form(f) => {
                        format!("{} Selected | Form Designer", f.selected_controls.len())
                    }
                    TabContent::Resources(r) => {
                        format!("{} resources | {}", r.entries.len(), r.active_tab.label())
                    }
                }
            } else {
                String::new()
            };
            self.status_bar
                .add_section(&status_text, status_text.len() as f32 * 7.5, false);

            // Right sections (right to left)
            let lang_label = format!("Language: {}", self.current_lang);
            self.status_bar.add_section_with_id(
                &lang_label,
                lang_label.len() as f32 * 7.5,
                true,
                "lang",
            );
            let theme_label = format!("Theme: {}", theme_name);
            self.status_bar.add_section_with_id(
                &theme_label,
                theme_label.len() as f32 * 7.5,
                true,
                "theme",
            );
            let config_label = match self.build_config {
                super::BuildConfig::Debug => "[Debug]",
                super::BuildConfig::Release => "[Release]",
            };
            self.status_bar.add_section_with_id(
                config_label,
                config_label.len() as f32 * 7.5,
                true,
                "config",
            );

            // Apply theme text color to all sections (default is white, wrong on light themes)
            let (ftr, ftg, ftb, fta) = (
                theme.footer_text.r(),
                theme.footer_text.g(),
                theme.footer_text.b(),
                theme.footer_text.a(),
            );
            for sec in self.status_bar.sections_mut() {
                sec.fg = (ftr, ftg, ftb, fta);
            }

            let mut ctx = RenderContext {
                pixmap: pix,
                font_system: &mut self.font_system,
                swash_cache: &mut self.swash_cache,
                scale: SCALE,
            };
            self.status_bar.render(&mut ctx);
        }

        // Dropdown overlays — positions in true logical pixels (physical / display_scale).
        // render_list(x,y) multiplies by dropdown.scale == display_scale → correct physical coords.
        // Hover/click handlers use self.win_width which is also physical / display_scale.
        {
            let win_w = self.win_width;
            let win_h = self.win_height;
            let menu_y_base = win_h - FOOTER_HEIGHT;
            let dd_colors = (
                (
                    theme.sidebar_bg.r(),
                    theme.sidebar_bg.g(),
                    theme.sidebar_bg.b(),
                    255u8,
                ),
                (
                    theme.gutter_divider.r(),
                    theme.gutter_divider.g(),
                    theme.gutter_divider.b(),
                    255u8,
                ),
                (
                    theme.selection.r(),
                    theme.selection.g(),
                    theme.selection.b(),
                    100u8,
                ),
                (
                    theme.current_line.r(),
                    theme.current_line.g(),
                    theme.current_line.b(),
                    255u8,
                ),
                theme.active_tab_text,
                theme.inactive_tab_text,
            );

            if let Some(dropdown) = &self.lang_dropdown {
                let (w, h) = dropdown.get_size();
                let menu_x = (win_w - w - 10.0).max(10.0);
                let menu_y = (menu_y_base - h - 5.0).max(0.0);
                dropdown.render_list(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    menu_x,
                    menu_y,
                    dd_colors.0,
                    dd_colors.1,
                    dd_colors.2,
                    dd_colors.3,
                    dd_colors.4,
                    dd_colors.5,
                );
            }
            if let Some(dropdown) = &self.theme_dropdown {
                let (w, h) = dropdown.get_size();
                // Anchor theme dropdown just to the left of the language dropdown
                let lang_w = format!("Language: {}", self.current_lang).len() as f32 * 7.5 + 16.0;
                let menu_x = (win_w - lang_w - w - 10.0).max(10.0);
                let menu_y = (menu_y_base - h - 5.0).max(0.0);
                dropdown.render_list(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    menu_x,
                    menu_y,
                    dd_colors.0,
                    dd_colors.1,
                    dd_colors.2,
                    dd_colors.3,
                    dd_colors.4,
                    dd_colors.5,
                );
            }
        }

        // 5. Diagnostic Tooltip on Hover
        if let Some(_tab) = self.tabs.get(self.active_tab) {
            let _mx = self.mouse_pos.0;
            let _my = self.mouse_pos.1;
        }

        // Quick Open Overlay
        if self.is_quick_open {
            let mut o_p = Paint::default();
            o_p.set_color_rgba8(30, 30, 35, 240);
            let o_w = 400.0 * SCALE;
            let o_h = 300.0 * SCALE;
            let o_x = (pix.width() as f32 - o_w) / 2.0;
            let o_y = 100.0 * SCALE;
            pix.fill_rect(
                Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap(),
                &o_p,
                Transform::identity(),
                None,
            );
            let mut b_p = Paint::default();
            b_p.set_color_rgba8(80, 80, 90, 255);
            let mut pb = PathBuilder::new();
            pb.push_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap());
            if let Some(path) = pb.finish() {
                pix.stroke_path(
                    &path,
                    &b_p,
                    &Stroke {
                        width: 1.0 * SCALE,
                        ..Default::default()
                    },
                    Transform::identity(),
                    None,
                );
            }
            App::draw_ui_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &format!("Go to file: {}|", self.quick_open_query),
                o_x + 10.0 * SCALE,
                o_y + 10.0 * SCALE,
                TextColor::rgb(200, 200, 200),
            );
            let matcher = SkimMatcherV2::default();
            let mut matches: Vec<(i64, usize, &String)> = self
                .tabs
                .iter()
                .enumerate()
                .filter_map(|(idx, tab)| {
                    if self.quick_open_query.is_empty() {
                        Some((0, idx, &tab.name))
                    } else {
                        matcher
                            .fuzzy_match(&tab.name, &self.quick_open_query)
                            .map(|score| (score, idx, &tab.name))
                    }
                })
                .collect();
            matches.sort_by_key(|m| -m.0);
            let mut i_y = o_y + 50.0 * SCALE;
            for (idx, (_score, _tab_idx, name)) in matches.iter().take(10).enumerate() {
                let col = if idx == 0 {
                    TextColor::rgb(0, 122, 204)
                } else {
                    TextColor::rgb(200, 200, 200)
                };
                let display_text = name.to_string();
                App::draw_ui_text(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    &display_text,
                    o_x + 20.0 * SCALE,
                    i_y,
                    col,
                );
                i_y += 25.0 * SCALE;
            }
        }

        // Go-to-Line Overlay
        if self.goto_line_open {
            let gl_w = 300.0 * SCALE;
            let gl_h = 50.0 * SCALE;
            let gl_x = (pix.width() as f32 - gl_w) / 2.0;
            let gl_y = 100.0 * SCALE;
            let mut bg = Paint::default();
            bg.set_color_rgba8(30, 30, 35, 245);
            pix.fill_rect(
                Rect::from_xywh(gl_x, gl_y, gl_w, gl_h).unwrap(),
                &bg,
                Transform::identity(),
                None,
            );
            let mut bp = Paint::default();
            bp.set_color_rgba8(0, 122, 204, 255);
            let mut pb = PathBuilder::new();
            pb.push_rect(Rect::from_xywh(gl_x, gl_y, gl_w, gl_h).unwrap());
            if let Some(path) = pb.finish() {
                pix.stroke_path(
                    &path,
                    &bp,
                    &Stroke {
                        width: SCALE,
                        ..Default::default()
                    },
                    Transform::identity(),
                    None,
                );
            }
            let max_line = self
                .tabs
                .get(self.active_tab)
                .map(|t| match &t.content {
                    TabContent::Code(cw) => cw.line_count(),
                    _ => 0,
                })
                .unwrap_or(0);
            App::draw_ui_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &format!("Go to Line (1-{}): {}|", max_line, self.goto_line_query),
                gl_x + 12.0 * SCALE,
                gl_y + 16.0 * SCALE,
                TextColor::rgb(200, 200, 200),
            );
        }

        // Command palette overlay
        if self.is_command_palette {
            let o_w = 520.0 * SCALE;
            let o_h = 420.0 * SCALE;
            let o_x = (pix.width() as f32 - o_w) / 2.0;
            let o_y = 90.0 * SCALE;
            let mut bg = Paint::default();
            bg.set_color_rgba8(30, 30, 35, 245);
            pix.fill_rect(
                Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap(),
                &bg,
                Transform::identity(),
                None,
            );
            let mut bp = Paint::default();
            bp.set_color_rgba8(80, 80, 90, 255);
            let mut pb = PathBuilder::new();
            pb.push_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap());
            if let Some(path) = pb.finish() {
                pix.stroke_path(
                    &path,
                    &bp,
                    &Stroke {
                        width: SCALE,
                        ..Default::default()
                    },
                    Transform::identity(),
                    None,
                );
            }
            App::draw_ui_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &format!("> {}|", self.command_palette_query),
                o_x + 12.0 * SCALE,
                o_y + 10.0 * SCALE,
                TextColor::rgb(220, 220, 220),
            );
            let matches = self.command_palette_matches();
            let cmds = super::palette_commands();
            let mut iy = o_y + 44.0 * SCALE;
            for (row, idx) in matches.iter().take(14).enumerate() {
                if row == self.command_palette_selected {
                    let mut hp = Paint::default();
                    hp.set_color_rgba8(0, 122, 204, 70);
                    pix.fill_rect(
                        Rect::from_xywh(o_x + 4.0, iy - 2.0, o_w - 8.0, 22.0 * SCALE).unwrap(),
                        &hp,
                        Transform::identity(),
                        None,
                    );
                }
                let col = if row == self.command_palette_selected {
                    TextColor::rgb(240, 240, 240)
                } else {
                    TextColor::rgb(190, 190, 190)
                };
                App::draw_ui_text(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    cmds[*idx].label,
                    o_x + 20.0 * SCALE,
                    iy,
                    col,
                );
                // Keybinding hint on the right
                let hint = kb_hint_for_palette(&self.keybindings, cmds[*idx].action);
                if !hint.is_empty() {
                    App::draw_ui_text(
                        pix,
                        &mut self.font_system,
                        &mut self.swash_cache,
                        &hint,
                        o_x + o_w - hint.len() as f32 * 8.0 * SCALE - 16.0,
                        iy,
                        TextColor::rgb(120, 120, 130),
                    );
                }
                iy += 22.0 * SCALE;
            }
            if matches.is_empty() {
                App::draw_ui_text(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    "(no matching commands)",
                    o_x + 20.0 * SCALE,
                    iy,
                    TextColor::rgb(120, 120, 120),
                );
            }
        }

        // Project-wide search overlay
        if self.is_project_search {
            let o_w = 640.0 * SCALE;
            let o_h = 460.0 * SCALE;
            let o_x = (pix.width() as f32 - o_w) / 2.0;
            let o_y = 90.0 * SCALE;
            let mut bg = Paint::default();
            bg.set_color_rgba8(30, 30, 35, 245);
            pix.fill_rect(
                Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap(),
                &bg,
                Transform::identity(),
                None,
            );
            let mut bp = Paint::default();
            bp.set_color_rgba8(80, 80, 90, 255);
            let mut pb = PathBuilder::new();
            pb.push_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap());
            if let Some(path) = pb.finish() {
                pix.stroke_path(
                    &path,
                    &bp,
                    &Stroke {
                        width: SCALE,
                        ..Default::default()
                    },
                    Transform::identity(),
                    None,
                );
            }
            App::draw_ui_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &format!("Find in Project: {}|", self.project_search_query),
                o_x + 12.0 * SCALE,
                o_y + 10.0 * SCALE,
                TextColor::rgb(220, 220, 220),
            );
            let count = self.project_search_results.len();
            let sub = if count == 0 && self.project_search_query.trim().len() >= 2 {
                "(no matches)".to_string()
            } else if count >= 500 {
                "500+ matches (narrow query)".to_string()
            } else {
                format!("{} match{}", count, if count == 1 { "" } else { "es" })
            };
            App::draw_ui_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &sub,
                o_x + 12.0 * SCALE,
                o_y + 32.0 * SCALE,
                TextColor::rgb(140, 140, 140),
            );

            let mut iy = o_y + 58.0 * SCALE;
            for (row, hit) in self.project_search_results.iter().take(16).enumerate() {
                if row == self.project_search_selected {
                    let mut hp = Paint::default();
                    hp.set_color_rgba8(0, 122, 204, 70);
                    pix.fill_rect(
                        Rect::from_xywh(o_x + 4.0, iy - 2.0, o_w - 8.0, 22.0 * SCALE).unwrap(),
                        &hp,
                        Transform::identity(),
                        None,
                    );
                }
                let label = format!("{}:{}  {}", hit.file, hit.line + 1, hit.snippet);
                let col = if row == self.project_search_selected {
                    TextColor::rgb(240, 240, 240)
                } else {
                    TextColor::rgb(190, 190, 190)
                };
                App::draw_ui_text(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    &label,
                    o_x + 16.0 * SCALE,
                    iy,
                    col,
                );
                iy += 22.0 * SCALE;
            }
        }

        // Tab context menu overlay
        if let Some((cmx, cmy, tab_idx)) = self.tab_context_menu {
            let entries = [
                "Close",
                "Close Others",
                "Close All",
                if self.tabs.get(tab_idx).map(|t| t.is_sticky).unwrap_or(false) {
                    "Unpin"
                } else {
                    "Pin"
                },
            ];
            let w = 180.0 * SCALE;
            let row_h = 26.0 * SCALE;
            let h = row_h * entries.len() as f32 + 8.0 * SCALE;
            let x = cmx * SCALE;
            let y = cmy * SCALE;
            let mut bg = Paint::default();
            bg.set_color_rgba8(35, 35, 42, 245);
            pix.fill_rect(
                Rect::from_xywh(x, y, w, h).unwrap(),
                &bg,
                Transform::identity(),
                None,
            );
            let mut bp = Paint::default();
            bp.set_color_rgba8(80, 80, 90, 255);
            let mut pb = PathBuilder::new();
            pb.push_rect(Rect::from_xywh(x, y, w, h).unwrap());
            if let Some(path) = pb.finish() {
                pix.stroke_path(
                    &path,
                    &bp,
                    &Stroke {
                        width: SCALE,
                        ..Default::default()
                    },
                    Transform::identity(),
                    None,
                );
            }
            for (i, e) in entries.iter().enumerate() {
                let row_y = y + 4.0 * SCALE + i as f32 * row_h;
                App::draw_ui_text(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    e,
                    x + 12.0 * SCALE,
                    row_y + 4.0 * SCALE,
                    TextColor::rgb(220, 220, 220),
                );
            }
        }

        // Menu dropdown overlay
        if has_form_tab {
            if let Some(form_tab) = self
                .tabs
                .iter()
                .find(|t| matches!(&t.content, TabContent::Form(_)))
            {
                if let TabContent::Form(f) = &form_tab.content {
                    let menu_rect = crate::form_designer_tab::Rect {
                        x: 0.0,
                        y: 0.0,
                        w: pix.width() as f32 / SCALE,
                        h: 28.0,
                    };
                    f.menu_bar.render_dropdown_overlay(
                        pix,
                        &mut self.font_system,
                        &mut self.swash_cache,
                        menu_rect,
                        SCALE,
                    );
                }
            }
        }

        // Project properties dialog (modal overlay)
        {
            let win_w = pix.width() as f32 / SCALE;
            let win_h = pix.height() as f32 / SCALE;
            self.project_props_dialog.render(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                win_w,
                win_h,
                SCALE,
                &self.project,
            );
        }

        // Project explorer context menu overlay
        if let Some((cmx, cmy, ref item_name)) = self.pe_context_menu {
            let mut cmp = Paint::default();
            let menu_w = 160.0f32;
            let menu_h = 28.0f32;
            cmp.set_color_rgba8(0, 0, 0, 40);
            if let Some(r) = Rect::from_xywh(
                (cmx + 2.0) * SCALE,
                (cmy + 2.0) * SCALE,
                menu_w * SCALE,
                menu_h * SCALE,
            ) {
                pix.fill_rect(r, &cmp, Transform::identity(), None);
            }
            cmp.set_color_rgba8(255, 255, 255, 255);
            if let Some(r) =
                Rect::from_xywh(cmx * SCALE, cmy * SCALE, menu_w * SCALE, menu_h * SCALE)
            {
                pix.fill_rect(r, &cmp, Transform::identity(), None);
            }
            cmp.set_color_rgba8(200, 200, 200, 255);
            let mut pb = PathBuilder::new();
            if let Some(r) =
                Rect::from_xywh(cmx * SCALE, cmy * SCALE, menu_w * SCALE, menu_h * SCALE)
            {
                pb.push_rect(r);
            }
            if let Some(path) = pb.finish() {
                let mut st = Stroke::default();
                st.width = SCALE;
                pix.stroke_path(&path, &cmp, &st, Transform::identity(), None);
            }
            let label = format!("\u{1F5D1} Remove \"{}\"", item_name);
            crate::ide_text::draw_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &label,
                cmx + 10.0,
                cmy + 6.0,
                12.0,
                TextColor::rgba(180, 40, 40, 255),
                SCALE,
            );
        }
    }

    /// Render the project explorer sidebar panel
    fn render_project_explorer(
        &mut self,
        pix: &mut Pixmap,
        sidebar_x: f32,
        sidebar_content_y: f32,
        _sidebar_w: f32,
        sidebar_content_h: f32,
        theme: &widgets::code_editor_widget::Theme,
    ) {
        let pe = &self.project_explorer;
        let project = &self.project;
        let current_form: Option<&str> = if self.active_tab < self.tabs.len() {
            if let TabContent::Form(f) = &self.tabs[self.active_tab].content {
                Some(&f.form.name)
            } else {
                None
            }
        } else {
            None
        };
        let pe_x = sidebar_x / SCALE;
        let pe_y = sidebar_content_y / SCALE;
        let pe_w = self.explorer_width;
        let pe_h = sidebar_content_h / SCALE;
        let item_h = 24.0f32;
        let indent = 16.0f32;
        let sel_bg = (
            theme.selection.r(),
            theme.selection.g(),
            theme.selection.b(),
            80u8,
        );
        let text_col = TextColor::rgba(
            theme.sidebar_text.r(),
            theme.sidebar_text.g(),
            theme.sidebar_text.b(),
            theme.sidebar_text.a(),
        );
        let dim_col = TextColor::rgba(
            theme.sidebar_text.r().saturating_sub(60),
            theme.sidebar_text.g().saturating_sub(60),
            theme.sidebar_text.b().saturating_sub(60),
            255,
        );
        let mut iy = pe_y - pe.scroll_y;
        let mut pp = Paint::default();

        // Project name
        if iy + item_h > pe_y && iy < pe_y + pe_h {
            crate::ide_text::draw_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &format!("\u{1F4C1} {}", project.name),
                pe_x + 8.0,
                iy + 4.0,
                12.0,
                text_col,
                SCALE,
            );
        }
        iy += item_h;

        // Forms section
        let forms_arrow = if pe.forms_collapsed {
            "\u{25B6}"
        } else {
            "\u{25BC}"
        };
        if iy + item_h > pe_y && iy < pe_y + pe_h {
            crate::ide_text::draw_text(
                pix,
                &mut self.font_system,
                &mut self.swash_cache,
                &format!("{} Forms", forms_arrow),
                pe_x + 8.0 + indent,
                iy + 4.0,
                12.0,
                text_col,
                SCALE,
            );
        }
        iy += item_h;

        if !pe.forms_collapsed {
            for fm in &project.forms {
                if iy + item_h > pe_y && iy < pe_y + pe_h {
                    let is_sel = current_form == Some(fm.form.name.as_str());
                    if is_sel {
                        pp.set_color_rgba8(sel_bg.0, sel_bg.1, sel_bg.2, sel_bg.3);
                        if let Some(r) = tiny_skia::Rect::from_xywh(
                            pe_x * SCALE,
                            iy * SCALE,
                            pe_w * SCALE,
                            item_h * SCALE,
                        ) {
                            pix.fill_rect(r, &pp, Transform::identity(), None);
                        }
                    }
                    crate::ide_text::draw_text(
                        pix,
                        &mut self.font_system,
                        &mut self.swash_cache,
                        &format!("  {}", fm.form.name),
                        pe_x + 8.0 + indent * 2.0,
                        iy + 4.0,
                        12.0,
                        text_col,
                        SCALE,
                    );
                }
                iy += item_h;
            }
        }

        // Code section
        if !project.code_files.is_empty() {
            let code_arrow = if pe.code_collapsed {
                "\u{25B6}"
            } else {
                "\u{25BC}"
            };
            if iy + item_h > pe_y && iy < pe_y + pe_h {
                crate::ide_text::draw_text(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    &format!("{} Code", code_arrow),
                    pe_x + 8.0 + indent,
                    iy + 4.0,
                    12.0,
                    text_col,
                    SCALE,
                );
            }
            iy += item_h;
            if !pe.code_collapsed {
                for cf in &project.code_files {
                    if iy + item_h > pe_y && iy < pe_y + pe_h {
                        crate::ide_text::draw_text(
                            pix,
                            &mut self.font_system,
                            &mut self.swash_cache,
                            &format!("  {}", cf.name),
                            pe_x + 8.0 + indent * 2.0,
                            iy + 4.0,
                            12.0,
                            text_col,
                            SCALE,
                        );
                    }
                    iy += item_h;
                }
            }
        }

        // References section
        if !project.project_references.is_empty() {
            let refs_arrow = if pe.refs_collapsed {
                "\u{25B6}"
            } else {
                "\u{25BC}"
            };
            if iy + item_h > pe_y && iy < pe_y + pe_h {
                crate::ide_text::draw_text(
                    pix,
                    &mut self.font_system,
                    &mut self.swash_cache,
                    &format!("{} References", refs_arrow),
                    pe_x + 8.0 + indent,
                    iy + 4.0,
                    12.0,
                    text_col,
                    SCALE,
                );
            }
            iy += item_h;
            if !pe.refs_collapsed {
                for rn in &project.project_references {
                    if iy + item_h > pe_y && iy < pe_y + pe_h {
                        crate::ide_text::draw_text(
                            pix,
                            &mut self.font_system,
                            &mut self.swash_cache,
                            &format!("  {}", rn),
                            pe_x + 8.0 + indent * 2.0,
                            iy + 4.0,
                            12.0,
                            dim_col,
                            SCALE,
                        );
                    }
                    iy += item_h;
                }
            }
        }

        // Resources section
        {
            let has_any_resources = !project.resource_files.is_empty()
                && project
                    .resource_files
                    .iter()
                    .any(|rm| !rm.resources.is_empty() || rm.file_path.is_some());
            let has_res_tab = self
                .tabs
                .iter()
                .any(|t| matches!(&t.content, TabContent::Resources(_)));
            if has_any_resources || has_res_tab {
                let res_count: usize = project
                    .resource_files
                    .iter()
                    .map(|rm| rm.resources.len())
                    .sum();
                let res_arrow = if pe.resources_collapsed {
                    "\u{25B6}"
                } else {
                    "\u{25BC}"
                };
                let res_label = if res_count > 0 {
                    format!("{} Resources ({})", res_arrow, res_count)
                } else {
                    format!("{} Resources", res_arrow)
                };
                if iy + item_h > pe_y && iy < pe_y + pe_h {
                    crate::ide_text::draw_text(
                        pix,
                        &mut self.font_system,
                        &mut self.swash_cache,
                        &res_label,
                        pe_x + 8.0 + indent,
                        iy + 4.0,
                        12.0,
                        text_col,
                        SCALE,
                    );
                }
                iy += item_h;
                if !pe.resources_collapsed {
                    for (_ri, rm) in project.resource_files.iter().enumerate() {
                        let rm_label = format!("  {}.resx ({})", rm.name, rm.resources.len());
                        let is_res_tab = self
                            .tabs
                            .get(self.active_tab)
                            .map(|t| matches!(&t.content, TabContent::Resources(_)))
                            .unwrap_or(false);
                        if iy + item_h > pe_y && iy < pe_y + pe_h {
                            if is_res_tab {
                                pp.set_color_rgba8(sel_bg.0, sel_bg.1, sel_bg.2, sel_bg.3);
                                if let Some(r) = tiny_skia::Rect::from_xywh(
                                    pe_x * SCALE,
                                    iy * SCALE,
                                    pe_w * SCALE,
                                    item_h * SCALE,
                                ) {
                                    pix.fill_rect(r, &pp, Transform::identity(), None);
                                }
                            }
                            crate::ide_text::draw_text(
                                pix,
                                &mut self.font_system,
                                &mut self.swash_cache,
                                &rm_label,
                                pe_x + 8.0 + indent * 2.0,
                                iy + 4.0,
                                12.0,
                                text_col,
                                SCALE,
                            );
                        }
                        iy += item_h;
                    }
                }
            }
        }
    }
}
