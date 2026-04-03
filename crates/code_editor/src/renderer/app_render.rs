use cosmic_text::{Attrs, Buffer, Color, Family, Metrics, Shaping};
use fuzzy_matcher::FuzzyMatcher;
use fuzzy_matcher::skim::SkimMatcherV2;
use tiny_skia::{Color as SkiaColor, Paint, Pixmap, PixmapPaint, Rect, Transform, ColorU8, Stroke, PathBuilder};

use cosmic_text::Edit;
use super::{App, TabContent, SidebarTab, BottomPanelTab, SCALE, TAB_BAR_HEIGHT, MINIMAP_WIDTH, UI_BAR_HEIGHT, FOOTER_HEIGHT, GUTTER_WIDTH, SPLITTER_WIDTH, SIDEBAR_TAB_H};
use crate::lsp_client::{LspRequest, LspEvent};

impl App {
    pub(super) fn render_internal(&mut self, pix: &mut Pixmap) {
        // Debounce LSP Update
        if self.pending_lsp_update && self.last_lsp_update.elapsed().as_millis() > 300 {
            let mut lsp_text = None;
            if let Some(tab) = self.tabs.get(self.active_tab) {
                if let TabContent::Code(cw) = &tab.content {
                    let text = cw.my_editor.rope.to_string();
                    let uri = tab.path.clone().unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", tab.name));
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
        
        let to_skia = |c: Color| SkiaColor::from_rgba8(c.r(), c.g(), c.b(), c.a());
        pix.fill(to_skia(theme.bg));

        // Compute top chrome height (menu + toolbar when any Form tab exists)
        let has_form_tab = self.tabs.iter().any(|t| matches!(&t.content, TabContent::Form(_)));
        let top_chrome_h: f32 = if has_form_tab { 28.0 + 36.0 } else { 0.0 };
        let top_chrome_px = top_chrome_h * SCALE;

        // 0. Menu bar + Toolbar (always present when a Form tab exists)
        if has_form_tab {
            if let Some(form_tab) = self.tabs.iter().find(|t| matches!(&t.content, TabContent::Form(_))) {
                if let TabContent::Form(f) = &form_tab.content {
                    let menu_rect = crate::form_designer_tab::Rect { x: 0.0, y: 0.0, w: pix.width() as f32 / SCALE, h: 28.0 };
                    let tb_rect = crate::form_designer_tab::Rect { x: 0.0, y: 28.0, w: pix.width() as f32 / SCALE, h: 36.0 };
                    f.menu_bar.render(pix, &mut self.font_system, &mut self.swash_cache, menu_rect, SCALE);
                    crate::form_designer_tab::render_toolbar_pub(pix, &mut self.font_system, &mut self.swash_cache, tb_rect, SCALE);
                }
            }
        }

        // 1. Sidebar
        let sidebar_x = 0.0;
        let sidebar_top = top_chrome_px;
        let sidebar_w = self.explorer_width * SCALE;
        let sidebar_h = pix.height() as f32 - top_chrome_px;

        // Sidebar background
        let mut sp = Paint::default(); sp.set_color_rgba8(theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), theme.sidebar_bg.a());
        pix.fill_rect(Rect::from_xywh(sidebar_x, sidebar_top, sidebar_w, sidebar_h).unwrap(), &sp, Transform::identity(), None);

        // Sidebar tabs (Files | Project)
        let stab_h = SIDEBAR_TAB_H * SCALE;
        let stab_w = sidebar_w / 2.0;
        let stab_y = sidebar_top;
        // Files tab
        {
            let active = self.sidebar_tab == SidebarTab::Files;
            let mut tp = Paint::default();
            if active { tp.set_color_rgba8(theme.active_tab_bg.r(), theme.active_tab_bg.g(), theme.active_tab_bg.b(), theme.active_tab_bg.a()); }
            else { tp.set_color_rgba8(theme.inactive_tab_bg.r(), theme.inactive_tab_bg.g(), theme.inactive_tab_bg.b(), theme.inactive_tab_bg.a()); }
            pix.fill_rect(Rect::from_xywh(sidebar_x, stab_y, stab_w, stab_h).unwrap(), &tp, Transform::identity(), None);
            if active {
                let mut up = Paint::default(); up.set_color_rgba8(theme.kw.r(), theme.kw.g(), theme.kw.b(), 255);
                pix.fill_rect(Rect::from_xywh(sidebar_x, stab_y + stab_h - 2.0 * SCALE, stab_w, 2.0 * SCALE).unwrap(), &up, Transform::identity(), None);
            }
            let col = if active { theme.active_tab_text } else { theme.inactive_tab_text };
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "Files", sidebar_x + 10.0 * SCALE, stab_y + 6.0 * SCALE, col);
        }
        // Project tab
        {
            let active = self.sidebar_tab == SidebarTab::Project;
            let mut tp = Paint::default();
            if active { tp.set_color_rgba8(theme.active_tab_bg.r(), theme.active_tab_bg.g(), theme.active_tab_bg.b(), theme.active_tab_bg.a()); }
            else { tp.set_color_rgba8(theme.inactive_tab_bg.r(), theme.inactive_tab_bg.g(), theme.inactive_tab_bg.b(), theme.inactive_tab_bg.a()); }
            pix.fill_rect(Rect::from_xywh(sidebar_x + stab_w, stab_y, stab_w, stab_h).unwrap(), &tp, Transform::identity(), None);
            if active {
                let mut up = Paint::default(); up.set_color_rgba8(theme.kw.r(), theme.kw.g(), theme.kw.b(), 255);
                pix.fill_rect(Rect::from_xywh(sidebar_x + stab_w, stab_y + stab_h - 2.0 * SCALE, stab_w, 2.0 * SCALE).unwrap(), &up, Transform::identity(), None);
            }
            let col = if active { theme.active_tab_text } else { theme.inactive_tab_text };
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "Project", sidebar_x + stab_w + 10.0 * SCALE, stab_y + 6.0 * SCALE, col);
        }
        // Tab separator
        {
            let mut lp = Paint::default(); lp.set_color_rgba8(theme.splitter_bg.r(), theme.splitter_bg.g(), theme.splitter_bg.b(), 255);
            pix.fill_rect(Rect::from_xywh(sidebar_x, stab_y + stab_h - 1.0, sidebar_w, 1.0).unwrap(), &lp, Transform::identity(), None);
        }

        // Sidebar content below tabs
        let sidebar_content_y = stab_y + stab_h;
        match self.sidebar_tab {
            SidebarTab::Files => {
                if let Some(tab) = self.tabs.get(self.active_tab) {
                    if let Some(path) = &tab.path {
                        self.tree_view.reveal_path(path);
                    }
                }
                self.tree_view.render(pix, &mut self.font_system, &mut self.swash_cache, sidebar_x, sidebar_content_y, sidebar_w, theme.sidebar_text, (theme.selection.r(), theme.selection.g(), theme.selection.b(), theme.selection.a()));
            }
            SidebarTab::Project => {
                self.render_project_explorer(pix, sidebar_x, sidebar_content_y, sidebar_w, sidebar_h - stab_h, &theme);
            }
        }

        // 1b. Splitter
        let mut slp = Paint::default();
        if self.is_dragging_splitter { slp.set_color_rgba8(0,122,204,255); }
        else if self.hovering_splitter { slp.set_color_rgba8(theme.splitter_bg.r(), theme.splitter_bg.g(), theme.splitter_bg.b(), 255); }
        else { slp.set_color_rgba8(theme.splitter_bg.r(), theme.splitter_bg.g(), theme.splitter_bg.b(), 255); }
        pix.fill_rect(Rect::from_xywh(self.explorer_width * SCALE, top_chrome_px, SPLITTER_WIDTH * SCALE, pix.height() as f32 - top_chrome_px).unwrap(), &slp, Transform::identity(), None);

        let mut lp = Paint::default(); lp.set_color_rgba8(theme.splitter_bg.r().saturating_add(20), theme.splitter_bg.g().saturating_add(20), theme.splitter_bg.b().saturating_add(20), 255);
        pix.fill_rect(Rect::from_xywh((self.explorer_width + SPLITTER_WIDTH) * SCALE, top_chrome_px, 1.0 * SCALE, pix.height() as f32 - top_chrome_px).unwrap(), &lp, Transform::identity(), None);

        // 2. Tab Bar
        let ed_start_x = (self.explorer_width + SPLITTER_WIDTH + 1.0) * SCALE;
        let mut tp = Paint::default(); tp.set_color_rgba8(theme.tab_bar_bg.r(), theme.tab_bar_bg.g(), theme.tab_bar_bg.b(), theme.tab_bar_bg.a());
        pix.fill_rect(Rect::from_xywh(ed_start_x, top_chrome_px, pix.width() as f32 - ed_start_x, TAB_BAR_HEIGHT * SCALE).unwrap(), &tp, Transform::identity(), None);

            let mut tx_off = ed_start_x + self.tab_scroll_x;
            for i in 0..self.tabs.len() {
                if tx_off + 160.0 * SCALE < ed_start_x { tx_off += 160.0 * SCALE; continue; }
                if tx_off > pix.width() as f32 { break; }
                
                let active = i == self.active_tab;
                let tw = 160.0 * SCALE;
                
                if active {
                    let mut ap = Paint::default(); ap.set_color_rgba8(theme.active_tab_bg.r(), theme.active_tab_bg.g(), theme.active_tab_bg.b(), theme.active_tab_bg.a());
                    pix.fill_rect(Rect::from_xywh(tx_off, top_chrome_px, tw, TAB_BAR_HEIGHT * SCALE).unwrap(), &ap, Transform::identity(), None);
                    let mut up = Paint::default(); up.set_color_rgba8(theme.kw.r(), theme.kw.g(), theme.kw.b(), 255);
                    pix.fill_rect(Rect::from_xywh(tx_off, top_chrome_px + (TAB_BAR_HEIGHT - 2.0) * SCALE, tw, 2.0 * SCALE).unwrap(), &up, Transform::identity(), None);
                } else {
                    let mut ip = Paint::default(); ip.set_color_rgba8(theme.inactive_tab_bg.r(), theme.inactive_tab_bg.g(), theme.inactive_tab_bg.b(), theme.inactive_tab_bg.a());
                    pix.fill_rect(Rect::from_xywh(tx_off, top_chrome_px, tw, TAB_BAR_HEIGHT * SCALE).unwrap(), &ip, Transform::identity(), None);
                }

                let (is_sticky, name, is_modified) = {
                    let t = &self.tabs[i];
                    (t.is_sticky, t.name.clone(), t.is_modified)
                };
                let name_str = if is_sticky { name } else { format!("{} [P]", name) };
                let col = if active { theme.active_tab_text } else { theme.inactive_tab_text };

                let tab_mut = &mut self.tabs[i];
                if tab_mut.buffer.is_none() {
                    let mut lab = Buffer::new(&mut self.font_system, Metrics::new(14.0,20.0).scale(SCALE));
                    lab.set_text(&mut self.font_system, &name_str, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None);
                    lab.shape_until_scroll(&mut self.font_system, false);
                    tab_mut.buffer = Some(lab);
                }
                if let Some(lab) = &tab_mut.buffer {
                    for r in lab.layout_runs() {
                        for g in r.glyphs {
                            let pg = g.physical((tx_off + 10.0 * SCALE, top_chrome_px + 10.0 * SCALE + r.line_y), 1.0);
                            if let Some(im) = self.swash_cache.get_image(&mut self.font_system, pg.cache_key) {
                                let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap();
                                let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a());
                                for (idx, &al) in im.data.iter().enumerate() {
                                    let af = (al as f32 / 255.0) * (ca as f32 / 255.0);
                                    p.pixels_mut()[idx] = ColorU8::from_rgba((cr as f32 * af) as u8, (cg as f32 * af) as u8, (cb as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                                }
                                pix.draw_pixmap(pg.x + im.placement.left, pg.y - im.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                            }
                        }
                    }
                }
                
                if is_modified {
                    App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "•", tx_off + tw - 24.0 * SCALE, top_chrome_px + 10.0 * SCALE, Color::rgb(180, 180, 180));
                } else {
                    let is_close_hover = self.hovering_tab_close == Some(i);
                    App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "×", tx_off + tw - 24.0 * SCALE, top_chrome_px + 10.0 * SCALE, if is_close_hover { Color::rgb(255, 100, 100) } else { Color::rgb(120,120,120) });
                }

                tx_off += tw;
            }

        // 3. Active Editor or Designer
        let output_h = if self.output_visible { self.output_panel_height } else { 0.0 };
        if self.active_tab < self.tabs.len() {
             let ed_top = top_chrome_px + (TAB_BAR_HEIGHT + UI_BAR_HEIGHT) * SCALE;
             let rect = Rect::from_xywh(ed_start_x, ed_top, pix.width() as f32 - ed_start_x, pix.height() as f32 - (ed_top + (FOOTER_HEIGHT + output_h) * SCALE)).unwrap();

             // Drain LSP events
             while let Ok(evt) = self.lsp.rx.try_recv() {
                match evt {
                    LspEvent::Diagnostics(uri, diags) => { 
                        for t in &mut self.tabs {
                            let t_uri = t.path.clone().unwrap_or_else(|| format!("file:///Users/youness/www/html/vybe/{}", t.name));
                            if t_uri == uri {
                                if let TabContent::Code(cw) = &mut t.content {
                                    cw.my_editor.diagnostics = diags.iter().map(|d| {
                                        vybe_widgets::DiagnosticInfo {
                                            line: d.range.start.line as usize,
                                            col_start: d.range.start.character as usize,
                                            col_end: d.range.end.character as usize,
                                            message: d.message.clone(),
                                            severity: match d.severity {
                                                Some(lsp_types::DiagnosticSeverity::ERROR) => vybe_widgets::DiagnosticSeverity::Error,
                                                Some(lsp_types::DiagnosticSeverity::WARNING) => vybe_widgets::DiagnosticSeverity::Warning,
                                                Some(lsp_types::DiagnosticSeverity::INFORMATION) => vybe_widgets::DiagnosticSeverity::Info,
                                                _ => vybe_widgets::DiagnosticSeverity::Hint,
                                            },
                                        }
                                    }).collect();
                                    self.needs_redraw = true;
                                    break;
                                }
                            }
                        }
                    }
                    LspEvent::Completion(items) => {
                        use lsp_types::CompletionItemKind;
                        use vybe_widgets::code_editor_widget::AutocompleteItem;
                        if let Some(tab) = self.tabs.get_mut(self.active_tab) {
                            if let TabContent::Code(cw) = &mut tab.content {
                                cw.autocomplete_items = items.into_iter().map(|ci| {
                                    let kind_icon = match ci.kind {
                                        Some(CompletionItemKind::FUNCTION) | Some(CompletionItemKind::METHOD) => "fn",
                                        Some(CompletionItemKind::STRUCT) | Some(CompletionItemKind::CLASS) | Some(CompletionItemKind::INTERFACE) => "ty",
                                        Some(CompletionItemKind::KEYWORD) => "kw",
                                        Some(CompletionItemKind::MODULE) => "md",
                                        Some(CompletionItemKind::FIELD) | Some(CompletionItemKind::PROPERTY) => "fi",
                                        Some(CompletionItemKind::VARIABLE) => "va",
                                        Some(CompletionItemKind::CONSTANT) | Some(CompletionItemKind::ENUM_MEMBER) => "ct",
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
                                }).collect();
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
                            .file_name().map(|n| n.to_string_lossy().to_string())
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
                                cw.editor.set_cursor(cosmic_text::Cursor::new(pos.line as usize, pos.character as usize));
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
                     f.render(pix, &mut self.font_system, &mut self.swash_cache, crate::form_designer_tab::Rect { x: rect.left() / SCALE, y: rect.top() / SCALE, w: rect.width() / SCALE, h: rect.height() / SCALE }, SCALE);
                 }
                 TabContent::Code(w) => {
                     let wrap_lines = w.wrap_lines;
                     w.editor.with_buffer_mut(|b| {
                         let wrap = if wrap_lines { cosmic_text::Wrap::Word } else { cosmic_text::Wrap::None };
                         if b.wrap() != wrap {
                             b.set_wrap(&mut self.font_system, wrap);
                         }
                         if wrap_lines {
                             b.set_size(&mut self.font_system, Some(rect.width() - (GUTTER_WIDTH + MINIMAP_WIDTH) * SCALE), Some(rect.height()));
                         } else {
                             b.set_size(&mut self.font_system, Some(999999.0), Some(999999.0));
                         }
                     });
                     w.needs_reshape = true;
                     w.render(pix, &mut self.font_system, &mut self.swash_cache, rect);
                 }
                 TabContent::Resources(r) => {
                     r.render(pix, &mut self.font_system, &mut self.swash_cache, rect.left() / SCALE, rect.top() / SCALE, rect.width() / SCALE, rect.height() / SCALE, SCALE);
                 }
             }
        }

        // 3b. Output / Problems Panel (above footer)
        if self.output_visible {
            let out_x = ed_start_x;
            let out_h = self.output_panel_height * SCALE;
            let out_y = pix.height() as f32 - (FOOTER_HEIGHT + self.output_panel_height) * SCALE;
            let out_w = pix.width() as f32 - out_x;
            // Background
            let mut obg = Paint::default(); obg.set_color_rgba8(25, 25, 30, 255);
            if let Some(r) = Rect::from_xywh(out_x, out_y, out_w, out_h) { pix.fill_rect(r, &obg, Transform::identity(), None); }
            // Header bar
            let mut ohd = Paint::default(); ohd.set_color_rgba8(35, 35, 42, 255);
            if let Some(r) = Rect::from_xywh(out_x, out_y, out_w, 24.0 * SCALE) { pix.fill_rect(r, &ohd, Transform::identity(), None); }

            // Tab buttons: Output | Problems
            let tab_y = out_y + 4.0 * SCALE;
            let output_tab_x = out_x + 10.0 * SCALE;
            let output_col = if self.bottom_panel_tab == BottomPanelTab::Output { Color::rgb(230, 230, 230) } else { Color::rgb(120, 120, 120) };
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "Output", output_tab_x, tab_y, output_col);
            // Active tab underline for Output
            if self.bottom_panel_tab == BottomPanelTab::Output {
                let mut ul = Paint::default(); ul.set_color_rgba8(0, 122, 204, 255);
                if let Some(r) = Rect::from_xywh(output_tab_x, out_y + 22.0 * SCALE, 50.0 * SCALE, 2.0 * SCALE) { pix.fill_rect(r, &ul, Transform::identity(), None); }
            }
            let problems_tab_x = out_x + 80.0 * SCALE;
            // Count problems across all tabs
            let problem_count: usize = self.tabs.iter().map(|t| {
                if let TabContent::Code(cw) = &t.content { cw.my_editor.diagnostics.len() } else { 0 }
            }).sum();
            let problems_label = format!("Problems ({})", problem_count);
            let problems_col = if self.bottom_panel_tab == BottomPanelTab::Problems { Color::rgb(230, 230, 230) } else { Color::rgb(120, 120, 120) };
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &problems_label, problems_tab_x, tab_y, problems_col);
            if self.bottom_panel_tab == BottomPanelTab::Problems {
                let mut ul = Paint::default(); ul.set_color_rgba8(0, 122, 204, 255);
                if let Some(r) = Rect::from_xywh(problems_tab_x, out_y + 22.0 * SCALE, 100.0 * SCALE, 2.0 * SCALE) { pix.fill_rect(r, &ul, Transform::identity(), None); }
            }

            // Close button
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "×", out_x + out_w - 24.0 * SCALE, tab_y, Color::rgb(150, 150, 150));
            // Clear button (only for Output tab)
            if self.bottom_panel_tab == BottomPanelTab::Output {
                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "Clear", out_x + out_w - 80.0 * SCALE, tab_y, Color::rgb(120, 120, 120));
            }
            // Separator line
            let mut sep_p = Paint::default(); sep_p.set_color_rgba8(60, 60, 70, 255);
            if let Some(r) = Rect::from_xywh(out_x, out_y, out_w, 1.0) { pix.fill_rect(r, &sep_p, Transform::identity(), None); }

            let content_y = out_y + 24.0 * SCALE;
            let content_h = out_h - 24.0 * SCALE;
            let line_h = 18.0 * SCALE;
            let visible_lines = (content_h / line_h) as usize;

            match self.bottom_panel_tab {
                BottomPanelTab::Output => {
                    // Output lines
                    let skip = (self.output_scroll_y / 18.0).max(0.0) as usize;
                    for (i, line) in self.output_lines.iter().skip(skip).take(visible_lines + 1).enumerate() {
                        let ly = content_y + (i as f32) * line_h - (self.output_scroll_y % 18.0) * SCALE;
                        if ly >= content_y && ly < out_y + out_h {
                            let col = if line.starts_with("ERR:") || line.starts_with("Save error") {
                                Color::rgb(255, 100, 100)
                            } else if line.starts_with("Building") || line.starts_with("Running") {
                                Color::rgb(100, 200, 100)
                            } else {
                                Color::rgb(180, 180, 180)
                            };
                            let display = if line.len() > 120 { &line[..120] } else { line.as_str() };
                            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, display, out_x + 10.0 * SCALE, ly + 2.0 * SCALE, col);
                        }
                    }
                }
                BottomPanelTab::Problems => {
                    // Collect all diagnostics with file names
                    let all_diags: Vec<(String, usize, vybe_widgets::DiagnosticSeverity, String)> = self.tabs.iter().flat_map(|t| {
                        let name = t.name.clone();
                        if let TabContent::Code(cw) = &t.content {
                            cw.my_editor.diagnostics.iter().map(move |d| {
                                (name.clone(), d.line + 1, d.severity, d.message.clone())
                            }).collect::<Vec<_>>()
                        } else { vec![] }
                    }).collect();

                    if all_diags.is_empty() {
                        App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, "No problems detected.", out_x + 10.0 * SCALE, content_y + 4.0 * SCALE, Color::rgb(100, 200, 100));
                    } else {
                        let skip = (self.output_scroll_y / 18.0).max(0.0) as usize;
                        for (i, (file, line, severity, msg)) in all_diags.iter().skip(skip).take(visible_lines + 1).enumerate() {
                            let ly = content_y + (i as f32) * line_h - (self.output_scroll_y % 18.0) * SCALE;
                            if ly >= content_y && ly < out_y + out_h {
                                let (icon, icon_col) = match severity {
                                    vybe_widgets::DiagnosticSeverity::Error => ("●", Color::rgb(255, 80, 80)),
                                    vybe_widgets::DiagnosticSeverity::Warning => ("▲", Color::rgb(255, 200, 50)),
                                    vybe_widgets::DiagnosticSeverity::Info => ("ℹ", Color::rgb(80, 160, 255)),
                                    vybe_widgets::DiagnosticSeverity::Hint => ("…", Color::rgb(140, 140, 140)),
                                };
                                // Icon
                                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, icon, out_x + 10.0 * SCALE, ly + 2.0 * SCALE, icon_col);
                                // File:line
                                let loc = format!("{}:{}", file, line);
                                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &loc, out_x + 24.0 * SCALE, ly + 2.0 * SCALE, Color::rgb(130, 180, 230));
                                // Message (offset after file:line column)
                                let msg_display = if msg.len() > 100 { &msg[..100] } else { msg.as_str() };
                                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, msg_display, out_x + 200.0 * SCALE, ly + 2.0 * SCALE, Color::rgb(200, 200, 200));
                            }
                        }
                    }
                }
            }
        }

        // 4. Footer
        let mut fp = Paint::default(); fp.set_color_rgba8(theme.footer_bg.r(), theme.footer_bg.g(), theme.footer_bg.b(), theme.footer_bg.a());
        pix.fill_rect(Rect::from_xywh(0.0, pix.height() as f32 - FOOTER_HEIGHT * SCALE, pix.width() as f32, FOOTER_HEIGHT * SCALE).unwrap(), &fp, Transform::identity(), None);
        
        if let Some(tab) = self.tabs.get(self.active_tab) {
            let path_str = tab.path.clone().unwrap_or_else(|| tab.name.clone());
            let segments: Vec<&str> = path_str.split(|c| c == '/' || c == '\\').filter(|s| !s.is_empty()).collect();
            
            let status_prefix = match &tab.content {
                TabContent::Code(cw) => {
                    let cursor = cw.editor.cursor();
                    let text = cw.my_editor.rope.to_string();
                    let line_endings = if text.contains("\r\n") { "CRLF" } else { "LF" };
                    let zoom_pct = (cw.font_size / 14.0 * 100.0) as i32;
                    format!("Ln {}, Col {} | {}% | {} | UTF-8 | ", cursor.line + 1, cursor.index + 1, zoom_pct, line_endings)
                }
                TabContent::Form(f) => {
                    let sels = f.selected_controls.len();
                    format!("{} Selected | Form Designer | ", sels)
                }
                TabContent::Resources(r) => {
                    format!("{} resources | {} | ", r.entries.len(), r.active_tab.label())
                }
            };
            
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &status_prefix, 10.0 * SCALE, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
            
            let mut current_x = 10.0 * SCALE + (status_prefix.len() as f32 * 8.4 * SCALE);
            self.breadcrumb_rects.clear();
            for (i, seg) in segments.iter().enumerate() {
                let seg_text = if i == segments.len() - 1 { seg.to_string() } else { format!("{} > ", seg) };
                let seg_width = seg_text.len() as f32 * 8.4 * SCALE;
                let rect = Rect::from_xywh(current_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE, seg_width, FOOTER_HEIGHT * SCALE).unwrap();
                let partial_path = segments[0..=i].join("/");
                self.breadcrumb_rects.push((rect, partial_path));
                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &seg_text, current_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
                current_x += seg_width;
            }

            let lang_label = format!("Language: {}", self.current_lang);
            let theme_label = format!("Theme: {}", theme_name);
            let config_label = match self.build_config { super::BuildConfig::Debug => "Debug", super::BuildConfig::Release => "Release" };
            let label_x = pix.width() as f32 - (lang_label.len() as f32 * 9.0 + 20.0) * SCALE;
            let theme_x = label_x - (theme_label.len() as f32 * 9.0 + 30.0) * SCALE;
            let config_x = theme_x - ((config_label.len() + 2) as f32 * 9.0 + 20.0) * SCALE;
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &lang_label, label_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &theme_label, theme_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, theme.footer_text);
            // Build config indicator
            let config_col = match self.build_config { super::BuildConfig::Debug => Color::rgb(255, 180, 50), super::BuildConfig::Release => Color::rgb(100, 200, 100) };
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("[{}]", config_label), config_x, pix.height() as f32 - FOOTER_HEIGHT * SCALE + 4.0 * SCALE, config_col);

            if let Some(dropdown) = &self.lang_dropdown {
                let (w, h) = dropdown.get_size();
                let menu_x = (pix.width() as f32 / SCALE - w - 20.0).max(10.0);
                let menu_y = (pix.height() as f32 / SCALE - FOOTER_HEIGHT - h - 10.0).max(10.0);
                dropdown.render(pix, &mut self.font_system, &mut self.swash_cache, menu_x, menu_y,
                    (theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), 255),
                    (theme.gutter_divider.r(), theme.gutter_divider.g(), theme.gutter_divider.b(), 255),
                    (theme.selection.r(), theme.selection.g(), theme.selection.b(), 100),
                    (theme.current_line.r(), theme.current_line.g(), theme.current_line.b(), 255),
                    theme.active_tab_text, theme.inactive_tab_text);
            }
            if let Some(dropdown) = &self.theme_dropdown {
                let (w, h) = dropdown.get_size();
                let menu_x = (theme_x / SCALE - 10.0).max(10.0);
                let menu_x = menu_x.min(pix.width() as f32 / SCALE - w - 10.0).max(10.0);
                let menu_y = (pix.height() as f32 / SCALE - FOOTER_HEIGHT - h - 10.0).max(10.0);
                dropdown.render(pix, &mut self.font_system, &mut self.swash_cache, menu_x, menu_y,
                    (theme.sidebar_bg.r(), theme.sidebar_bg.g(), theme.sidebar_bg.b(), 255),
                    (theme.gutter_divider.r(), theme.gutter_divider.g(), theme.gutter_divider.b(), 255),
                    (theme.selection.r(), theme.selection.g(), theme.selection.b(), 100),
                    (theme.current_line.r(), theme.current_line.g(), theme.current_line.b(), 255),
                    theme.active_tab_text, theme.inactive_tab_text);
            }
        }

        // 5. Diagnostic Tooltip on Hover
        if let Some(_tab) = self.tabs.get(self.active_tab) {
            let _mx = self.mouse_pos.0; let _my = self.mouse_pos.1;
        }

        // Quick Open Overlay
        if self.is_quick_open {
            let mut o_p = Paint::default(); o_p.set_color_rgba8(30, 30, 35, 240);
            let o_w = 400.0 * SCALE; let o_h = 300.0 * SCALE;
            let o_x = (pix.width() as f32 - o_w) / 2.0; let o_y = 100.0 * SCALE;
            pix.fill_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap(), &o_p, Transform::identity(), None);
            let mut b_p = Paint::default(); b_p.set_color_rgba8(80, 80, 90, 255);
            let mut pb = PathBuilder::new(); pb.push_rect(Rect::from_xywh(o_x, o_y, o_w, o_h).unwrap());
            if let Some(path) = pb.finish() { pix.stroke_path(&path, &b_p, &Stroke { width: 1.0 * SCALE, ..Default::default() }, Transform::identity(), None); }
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("Go to file: {}|", self.quick_open_query), o_x + 10.0 * SCALE, o_y + 10.0 * SCALE, Color::rgb(200, 200, 200));
            let matcher = SkimMatcherV2::default();
            let mut matches: Vec<(i64, usize, &String)> = self.tabs.iter().enumerate()
                .filter_map(|(idx, tab)| {
                    if self.quick_open_query.is_empty() { Some((0, idx, &tab.name)) }
                    else { matcher.fuzzy_match(&tab.name, &self.quick_open_query).map(|score| (score, idx, &tab.name)) }
                }).collect();
            matches.sort_by_key(|m| -m.0);
            let mut i_y = o_y + 50.0 * SCALE;
            for (idx, (score, _tab_idx, name)) in matches.iter().take(10).enumerate() {
                let col = if idx == 0 { Color::rgb(0, 122, 204) } else { Color::rgb(200, 200, 200) };
                let display_text = if *score > 0 { format!("{} (score: {})", name, score) } else { name.to_string() };
                App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &display_text, o_x + 20.0 * SCALE, i_y, col);
                i_y += 25.0 * SCALE;
            }
        }

        // Go-to-Line Overlay
        if self.goto_line_open {
            let gl_w = 300.0 * SCALE; let gl_h = 50.0 * SCALE;
            let gl_x = (pix.width() as f32 - gl_w) / 2.0; let gl_y = 100.0 * SCALE;
            let mut bg = Paint::default(); bg.set_color_rgba8(30, 30, 35, 245);
            pix.fill_rect(Rect::from_xywh(gl_x, gl_y, gl_w, gl_h).unwrap(), &bg, Transform::identity(), None);
            let mut bp = Paint::default(); bp.set_color_rgba8(0, 122, 204, 255);
            let mut pb = PathBuilder::new(); pb.push_rect(Rect::from_xywh(gl_x, gl_y, gl_w, gl_h).unwrap());
            if let Some(path) = pb.finish() { pix.stroke_path(&path, &bp, &Stroke { width: SCALE, ..Default::default() }, Transform::identity(), None); }
            let max_line = self.tabs.get(self.active_tab).map(|t| match &t.content { TabContent::Code(cw) => cw.editor.with_buffer(|b| b.lines.len()), _ => 0 }).unwrap_or(0);
            App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("Go to Line (1-{}): {}|", max_line, self.goto_line_query), gl_x + 12.0 * SCALE, gl_y + 16.0 * SCALE, Color::rgb(200, 200, 200));
        }

        // Menu dropdown overlay
        if has_form_tab {
            if let Some(form_tab) = self.tabs.iter().find(|t| matches!(&t.content, TabContent::Form(_))) {
                if let TabContent::Form(f) = &form_tab.content {
                    let menu_rect = crate::form_designer_tab::Rect { x: 0.0, y: 0.0, w: pix.width() as f32 / SCALE, h: 28.0 };
                    f.menu_bar.render_dropdown_overlay(pix, &mut self.font_system, &mut self.swash_cache, menu_rect, SCALE);
                }
            }
        }

        // Project properties dialog (modal overlay)
        {
            let win_w = pix.width() as f32 / SCALE;
            let win_h = pix.height() as f32 / SCALE;
            self.project_props_dialog.render(pix, &mut self.font_system, &mut self.swash_cache, win_w, win_h, SCALE, &self.project);
        }

        // Project explorer context menu overlay
        if let Some((cmx, cmy, ref item_name)) = self.pe_context_menu {
            let mut cmp = Paint::default();
            let menu_w = 160.0f32; let menu_h = 28.0f32;
            cmp.set_color_rgba8(0, 0, 0, 40);
            if let Some(r) = Rect::from_xywh((cmx + 2.0) * SCALE, (cmy + 2.0) * SCALE, menu_w * SCALE, menu_h * SCALE) { pix.fill_rect(r, &cmp, Transform::identity(), None); }
            cmp.set_color_rgba8(255, 255, 255, 255);
            if let Some(r) = Rect::from_xywh(cmx * SCALE, cmy * SCALE, menu_w * SCALE, menu_h * SCALE) { pix.fill_rect(r, &cmp, Transform::identity(), None); }
            cmp.set_color_rgba8(200, 200, 200, 255);
            let mut pb = PathBuilder::new();
            if let Some(r) = Rect::from_xywh(cmx * SCALE, cmy * SCALE, menu_w * SCALE, menu_h * SCALE) { pb.push_rect(r); }
            if let Some(path) = pb.finish() { let mut st = Stroke::default(); st.width = SCALE; pix.stroke_path(&path, &cmp, &st, Transform::identity(), None); }
            let label = format!("\u{1F5D1} Remove \"{}\"", item_name);
            crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &label, cmx + 10.0, cmy + 6.0, 12.0, Color::rgba(180, 40, 40, 255), SCALE);
        }
    }

    /// Render the project explorer sidebar panel
    fn render_project_explorer(&mut self, pix: &mut Pixmap, sidebar_x: f32, sidebar_content_y: f32, _sidebar_w: f32, sidebar_content_h: f32, theme: &vybe_widgets::code_editor_widget::Theme) {
        let pe = &self.project_explorer;
        let project = &self.project;
        let current_form: Option<&str> = if self.active_tab < self.tabs.len() {
            if let TabContent::Form(f) = &self.tabs[self.active_tab].content { Some(&f.form.name) } else { None }
        } else { None };
        let pe_x = sidebar_x / SCALE;
        let pe_y = sidebar_content_y / SCALE;
        let pe_w = self.explorer_width;
        let pe_h = sidebar_content_h / SCALE;
        let item_h = 24.0f32;
        let indent = 16.0f32;
        let sel_bg = (theme.selection.r(), theme.selection.g(), theme.selection.b(), 80u8);
        let text_col = Color::rgba(theme.sidebar_text.r(), theme.sidebar_text.g(), theme.sidebar_text.b(), theme.sidebar_text.a());
        let dim_col = Color::rgba(theme.sidebar_text.r().saturating_sub(60), theme.sidebar_text.g().saturating_sub(60), theme.sidebar_text.b().saturating_sub(60), 255);
        let mut iy = pe_y - pe.scroll_y;
        let mut pp = Paint::default();

        // Project name
        if iy + item_h > pe_y && iy < pe_y + pe_h {
            crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("\u{1F4C1} {}", project.name), pe_x + 8.0, iy + 4.0, 12.0, text_col, SCALE);
        }
        iy += item_h;

        // Forms section
        let forms_arrow = if pe.forms_collapsed { "\u{25B6}" } else { "\u{25BC}" };
        if iy + item_h > pe_y && iy < pe_y + pe_h {
            crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("{} Forms", forms_arrow), pe_x + 8.0 + indent, iy + 4.0, 12.0, text_col, SCALE);
        }
        iy += item_h;

        if !pe.forms_collapsed {
            for fm in &project.forms {
                if iy + item_h > pe_y && iy < pe_y + pe_h {
                    let is_sel = current_form == Some(fm.form.name.as_str());
                    if is_sel {
                        pp.set_color_rgba8(sel_bg.0, sel_bg.1, sel_bg.2, sel_bg.3);
                        if let Some(r) = tiny_skia::Rect::from_xywh(pe_x * SCALE, iy * SCALE, pe_w * SCALE, item_h * SCALE) {
                            pix.fill_rect(r, &pp, Transform::identity(), None);
                        }
                    }
                    crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("  {}", fm.form.name), pe_x + 8.0 + indent * 2.0, iy + 4.0, 12.0, text_col, SCALE);
                }
                iy += item_h;
            }
        }

        // Code section
        if !project.code_files.is_empty() {
            let code_arrow = if pe.code_collapsed { "\u{25B6}" } else { "\u{25BC}" };
            if iy + item_h > pe_y && iy < pe_y + pe_h {
                crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("{} Code", code_arrow), pe_x + 8.0 + indent, iy + 4.0, 12.0, text_col, SCALE);
            }
            iy += item_h;
            if !pe.code_collapsed {
                for cf in &project.code_files {
                    if iy + item_h > pe_y && iy < pe_y + pe_h {
                        crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("  {}", cf.name), pe_x + 8.0 + indent * 2.0, iy + 4.0, 12.0, text_col, SCALE);
                    }
                    iy += item_h;
                }
            }
        }

        // References section
        if !project.project_references.is_empty() {
            let refs_arrow = if pe.refs_collapsed { "\u{25B6}" } else { "\u{25BC}" };
            if iy + item_h > pe_y && iy < pe_y + pe_h {
                crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("{} References", refs_arrow), pe_x + 8.0 + indent, iy + 4.0, 12.0, text_col, SCALE);
            }
            iy += item_h;
            if !pe.refs_collapsed {
                for rn in &project.project_references {
                    if iy + item_h > pe_y && iy < pe_y + pe_h {
                        crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("  {}", rn), pe_x + 8.0 + indent * 2.0, iy + 4.0, 12.0, dim_col, SCALE);
                    }
                    iy += item_h;
                }
            }
        }

        // Resources section
        {
            let has_any_resources = !project.resource_files.is_empty() &&
                project.resource_files.iter().any(|rm| !rm.resources.is_empty() || rm.file_path.is_some());
            let has_res_tab = self.tabs.iter().any(|t| matches!(&t.content, TabContent::Resources(_)));
            if has_any_resources || has_res_tab {
                let res_count: usize = project.resource_files.iter().map(|rm| rm.resources.len()).sum();
                let res_arrow = if pe.resources_collapsed { "\u{25B6}" } else { "\u{25BC}" };
                let res_label = if res_count > 0 { format!("{} Resources ({})", res_arrow, res_count) } else { format!("{} Resources", res_arrow) };
                if iy + item_h > pe_y && iy < pe_y + pe_h {
                    crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &res_label, pe_x + 8.0 + indent, iy + 4.0, 12.0, text_col, SCALE);
                }
                iy += item_h;
                if !pe.resources_collapsed {
                    for (_ri, rm) in project.resource_files.iter().enumerate() {
                        let rm_label = format!("  {}.resx ({})", rm.name, rm.resources.len());
                        let is_res_tab = self.tabs.get(self.active_tab).map(|t| matches!(&t.content, TabContent::Resources(_))).unwrap_or(false);
                        if iy + item_h > pe_y && iy < pe_y + pe_h {
                            if is_res_tab {
                                pp.set_color_rgba8(sel_bg.0, sel_bg.1, sel_bg.2, sel_bg.3);
                                if let Some(r) = tiny_skia::Rect::from_xywh(pe_x * SCALE, iy * SCALE, pe_w * SCALE, item_h * SCALE) {
                                    pix.fill_rect(r, &pp, Transform::identity(), None);
                                }
                            }
                            crate::ide_text::draw_text(pix, &mut self.font_system, &mut self.swash_cache, &rm_label, pe_x + 8.0 + indent * 2.0, iy + 4.0, 12.0, text_col, SCALE);
                        }
                        iy += item_h;
                    }
                }
            }
        }
    }
}
