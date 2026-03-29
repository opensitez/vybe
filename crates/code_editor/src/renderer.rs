use std::collections::{HashMap, HashSet};
use std::time::{Duration, Instant};
use std::num::NonZeroU32;
use std::sync::Arc;
use std::fs;
use cosmic_text::{Attrs, Buffer, Color, Family, FontSystem, Metrics, Motion, Shaping, SwashCache, Action, Edit, AttrsList};
use tiny_skia::{Color as SkiaColor, Paint, Pixmap, PixmapPaint, Rect, Transform, ColorU8};
use winit::application::ApplicationHandler;
use winit::event::{WindowEvent, ElementState, MouseButton, MouseScrollDelta};
use winit::event_loop::{ControlFlow, EventLoop};
use winit::window::{Window, WindowAttributes};
use winit::keyboard::{Key, NamedKey};
#[cfg(target_os = "macos")]
use winit::platform::modifier_supplement::KeyEventExtModifierSupplement;
use softbuffer::{Context, Surface};
use arboard::Clipboard;

use crate::editor::{Editor as MyEditor, TokenKind};
use crate::language::{load_language, LanguageDef};

const SCALE: f32 = 2.0;
const SIDEBAR_WIDTH: f32 = 70.0;
const TEXT_PADDING: f32 = 15.0;
const UI_BAR_HEIGHT: f32 = 40.0;

struct CachedGlyph {
    pixmap: Pixmap,
    left: i32,
    top: i32,
}

pub fn apply_highlighting(editor: &mut cosmic_text::Editor<'static>, my_editor: &MyEditor, attrs: &Attrs, lang: &LanguageDef) {
    let kw_color = Color::rgb(0x56, 0x9c, 0xd6);
    let type_kw_color = Color::rgb(0x4e, 0xc9, 0xb0);
    let str_color = Color::rgb(0xce, 0x91, 0x78);
    let comment_color = Color::rgb(0x6a, 0x99, 0x55);
    let num_color = Color::rgb(0xb5, 0xce, 0xa8);
    let ident_color = Color::rgb(0xee, 0xee, 0xee);

    editor.with_buffer_mut(|buffer| {
        let mut byte_offset = 0usize;
        for (li, line) in buffer.lines.iter_mut().enumerate() {
            let mut list = AttrsList::new(attrs);
            if li < my_editor.line_tokens.len() {
                for token in &my_editor.line_tokens[li] {
                    let color = match token.kind {
                        TokenKind::LineComment | TokenKind::BlockComment => comment_color,
                        TokenKind::String => str_color,
                        TokenKind::Number => num_color,
                        TokenKind::Punct => Color::rgb(0xd4, 0xd4, 0xd4),
                        TokenKind::Identifier => {
                            let text = my_editor.slice(token.start, token.end);
                            if lang.keywords.contains(&text) { kw_color }
                            else if lang.type_keywords.contains(&text) { type_kw_color }
                            else if lang.constants.contains(&text) { num_color }
                            else { ident_color }
                        }
                        _ => ident_color,
                    };
                    let start = token.start.saturating_sub(byte_offset);
                    let end = token.end.saturating_sub(byte_offset);
                    list.add_span(start..end, &attrs.clone().color(color));
                }
            }
            line.set_attrs_list(list);
            if li < my_editor.rope.len_lines() { byte_offset += my_editor.rope.line(li).len_bytes(); }
            else { byte_offset += line.text().len() + 1; }
        }
    });
}

pub struct CodeEditorWidget {
    editor: cosmic_text::Editor<'static>,
    my_editor: MyEditor,
    lang_def: LanguageDef,
    metrics: Metrics,
    glyph_cache: HashMap<(cosmic_text::CacheKey, Color), CachedGlyph>,
    digit_cache: Vec<CachedGlyph>,
    needs_reshape: bool,
    pub scroll_y: f32,
}

impl CodeEditorWidget {
    pub fn new(my_editor: MyEditor, font_system: &mut FontSystem) -> Self {
        let metrics = Metrics::new(14.0, 20.0).scale(SCALE);
        let lang_def = load_language("rust").unwrap_or_else(|| LanguageDef {
            keywords: HashSet::new(), type_keywords: HashSet::new(), constants: HashSet::new(), operators: HashSet::new(), comments: None, brackets: Vec::new(),
        });
        
        let mut editor_internal = my_editor;
        editor_internal.retokenize_all(&lang_def);
        let mut buffer = Buffer::new(font_system, metrics);
        buffer.set_text(font_system, &editor_internal.rope.to_string(), &Attrs::new().family(Family::Monospace), Shaping::Advanced, None);
        
        let digit_color = Color::rgb(0x85, 0x85, 0x85);
        let mut swash_cache = SwashCache::new();
        let mut digit_cache = Vec::new();
        for i in 0..10 {
            let mut lab = Buffer::new(font_system, metrics);
            lab.set_text(font_system, &format!("{}", i), &Attrs::new().family(Family::Monospace).color(digit_color), Shaping::Advanced, None);
            lab.shape_until_scroll(font_system, false);
            if let Some(r) = lab.layout_runs().next() {
                for g in r.glyphs {
                    let pg = g.physical((0.0, 0.0), 1.0);
                    if let Some(img) = swash_cache.get_image(font_system, pg.cache_key) {
                        let mut p = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)).unwrap();
                        let (r, g, b, a) = (digit_color.r(), digit_color.g(), digit_color.b(), digit_color.a());
                        for (idx, &alpha) in img.data.iter().enumerate() {
                            let af = (alpha as f32 / 255.0) * (a as f32 / 255.0);
                            p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                        }
                        digit_cache.push(CachedGlyph { pixmap: p, left: img.placement.left, top: img.placement.top });
                        break;
                    }
                }
            }
            if digit_cache.len() <= i { digit_cache.push(CachedGlyph { pixmap: Pixmap::new(1, 1).unwrap(), left: 0, top: 0 }); }
        }

        Self { editor: cosmic_text::Editor::new(buffer), my_editor: editor_internal, lang_def, metrics, glyph_cache: HashMap::new(), digit_cache, needs_reshape: true, scroll_y: 0.0 }
    }

    pub fn set_language(&mut self, lang_name: &str) {
        if let Some(lang) = load_language(lang_name) {
            self.lang_def = lang;
            self.my_editor.retokenize_all(&self.lang_def);
            self.needs_reshape = true;
        }
    }

    fn get_offsets(&self, rect: Rect) -> (f32, f32) {
        (rect.left() + SIDEBAR_WIDTH * SCALE, rect.top())
    }

    fn is_line_hidden(&self, line_idx: usize) -> bool {
        for (start, end) in &self.my_editor.folds {
            if self.my_editor.collapsed_starts.contains(start) {
                if line_idx > *start && line_idx <= *end { return true; }
            }
        }
        false
    }

    fn get_visual_y_shift(&self, line_idx: usize) -> f32 {
        let mut hidden_lines = 0;
        for (start, end) in &self.my_editor.folds {
            if self.my_editor.collapsed_starts.contains(start) {
                let current_hidden = end - start;
                if line_idx > *end { hidden_lines += current_hidden; }
            }
        }
        hidden_lines as f32 * self.metrics.line_height
    }

    pub fn render(&mut self, pixmap: &mut Pixmap, font_system: &mut FontSystem, swash_cache: &mut SwashCache, rect: Rect) {
        if self.needs_reshape {
            apply_highlighting(&mut self.editor, &self.my_editor, &Attrs::new().family(Family::Monospace), &self.lang_def);
            self.editor.shape_as_needed(font_system, false);
            self.needs_reshape = false;
        }

        let (x_off, y_off) = self.get_offsets(rect);
        let mut side_paint = Paint::default(); side_paint.set_color_rgba8(0x2d, 0x2d, 0x2d, 0xff);
        pixmap.fill_rect(Rect::from_xywh(rect.left(), rect.top(), SIDEBAR_WIDTH * SCALE, rect.height()).unwrap(), &side_paint, Transform::identity(), None);

        let cursor_state = self.editor.cursor();
        let selection = self.editor.selection_bounds();
        let mut selected_text = None;
        if let Some(t) = self.editor.copy_selection() { if t.len() > 1 && t.len() < 50 && !t.contains('\n') { selected_text = Some(t); } }

        // Find Partner Bracket
        let partner_bracket = self.my_editor.find_matching_bracket(cursor_state.line, cursor_state.index, &self.lang_def)
            .or_else(|| if cursor_state.index > 0 { self.my_editor.find_matching_bracket(cursor_state.line, cursor_state.index - 1, &self.lang_def) } else { None });

        // Auto-Scroll Logic
        if let Some((_, cy)) = self.editor.cursor_position() {
             let mut cursor_line_i = 0;
             self.editor.with_buffer(|b| {
                 for r in b.layout_runs() { if cy >= r.line_top as i32 && cy < (r.line_top + r.line_height) as i32 { cursor_line_i = r.line_i; break; } }
             });
             let cy_shifted = cy as f32 - self.get_visual_y_shift(cursor_line_i);
             let relative_cy = cy_shifted - self.scroll_y;
             if relative_cy < 0.0 { self.scroll_y = cy_shifted; }
             else if relative_cy > (rect.height() - self.metrics.line_height) { self.scroll_y = cy_shifted - (rect.height() - self.metrics.line_height); }
        }

        let mut runs_to_draw = Vec::new();
        self.editor.with_buffer(|buffer| {
            for run in buffer.layout_runs() {
                if !self.is_line_hidden(run.line_i) {
                    let y_shift = self.get_visual_y_shift(run.line_i);
                    let v_top = run.line_top - y_shift - self.scroll_y;
                    if v_top > rect.height() || (v_top + run.line_height) < 0.0 { continue; }
                    runs_to_draw.push((run.line_i, run.line_y, run.line_top, run.line_height, y_shift, run.glyphs.to_vec(), run.text.to_string()));
                }
            }
        });

        // 2. Highlighting Features
        // --- Current Line
        let mut cur_paint = Paint::default(); cur_paint.set_color_rgba8(0x32, 0x32, 0x32, 0xff);
        let cy_raw = self.editor.cursor_position().map(|(_, y)| y as f32).unwrap_or(-1000.0);
        let mut cursor_line_idx = 0;
        self.editor.with_buffer(|b| { for r in b.layout_runs() { if cy_raw >= (r.line_top as i32) as f32 && cy_raw < (r.line_top + r.line_height) as i32 as f32 { cursor_line_idx = r.line_i; } } });

        let mut last_para = None;
        for (line_i, line_y, line_top, line_height, y_shift, glyphs, line_text) in runs_to_draw {
            let current_y_off = y_off - y_shift - self.scroll_y;
            
            // --- Current Line Highlight
            if line_i == cursor_line_idx {
                pixmap.fill_rect(Rect::from_xywh(rect.left() + SIDEBAR_WIDTH * SCALE, current_y_off + line_top, rect.width() - SIDEBAR_WIDTH * SCALE, line_height).unwrap(), &cur_paint, Transform::identity(), None);
            }

            // --- Indentation Guides
            let mut guide_paint = Paint::default(); guide_paint.set_color_rgba8(0x40, 0x40, 0x40, 0xff);
            let tab_w = 4.0 * 8.4 * SCALE; // approximate monospace width
            let leading_spaces = line_text.chars().take_while(|c: &char| c.is_whitespace()).count();
            for i in 1..=(leading_spaces / 4) {
                let gx = x_off + (i as f32 * tab_w);
                pixmap.fill_rect(Rect::from_xywh(gx, current_y_off + line_top, 1.0, line_height).unwrap(), &guide_paint, Transform::identity(), None);
            }

            if last_para != Some(line_i) {
                let s = format!("{}", line_i + 1);
                let mut digit_x = (rect.left() + SIDEBAR_WIDTH * SCALE) as i32 - 15;
                for ch in s.chars().rev() {
                    if let Some(digit) = ch.to_digit(10) {
                        let cg = &self.digit_cache[digit as usize];
                        let y = (current_y_off + line_y) as i32 - cg.top;
                        pixmap.draw_pixmap(digit_x - cg.pixmap.width() as i32 + cg.left, y, cg.pixmap.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                        digit_x -= 10 * SCALE as i32;
                    }
                }
                if self.my_editor.folds.iter().any(|(s, _)| *s == line_i) {
                    let is_collapsed = self.my_editor.collapsed_starts.contains(&line_i);
                    let icon = if is_collapsed { "+" } else { "-" };
                    let color = if is_collapsed { Color::rgb(0x56, 0x9c, 0xd6) } else { Color::rgb(0x85, 0x85, 0x85) };
                    App::draw_ui_text(pixmap, font_system, swash_cache, icon, rect.left() + 5.0 * SCALE, current_y_off + line_y - 14.0 * SCALE, color);
                }
                last_para = Some(line_i);
            }

            // --- Selection Occurrences Highlight
            if let Some(target) = &selected_text {
                let mut start = 0;
                while let Some(idx) = line_text[start..].find(target) {
                    let absolute_idx = start + idx;
                    let mut head_x = None; let mut tail_x = 0.0;
                    for g in &glyphs {
                        if g.start >= absolute_idx && g.end <= absolute_idx + target.len() {
                            if head_x.is_none() { head_x = Some(g.x); }
                            tail_x = g.x + g.w;
                        }
                    }
                    if let Some(hx) = head_x {
                        let mut match_paint = Paint::default(); match_paint.set_color_rgba8(0x3e, 0x44, 0x51, 0xff);
                        pixmap.fill_rect(Rect::from_xywh(x_off + hx, current_y_off + line_top, tail_x - hx, line_height).unwrap(), &match_paint, Transform::identity(), None);
                    }
                    start += idx + target.len();
                }
            }

            if let Some((s_start, s_end)) = selection {
                self.editor.with_buffer(|buffer| {
                    if let Some(run) = buffer.layout_runs().nth(line_i) {
                        if let Some((hx, hw)) = run.highlight(s_start, s_end) {
                            let mut sp = Paint::default(); sp.set_color_rgba8(0x26, 0x4f, 0x78, 0xff);
                            pixmap.fill_rect(Rect::from_xywh(x_off + hx, current_y_off + line_top, hw, line_height).unwrap(), &sp, Transform::identity(), None);
                        }
                    }
                });
            }

            for glyph in &glyphs {
                let mut is_partner = false;
                if let Some((pl, pi)) = partner_bracket {
                    if line_i == pl && glyph.start == pi { is_partner = true; }
                }

                let pg = glyph.physical((x_off, current_y_off + line_y), 1.0);
                let gc = glyph.color_opt.unwrap_or(Color::rgb(0xee, 0xee, 0xee));
                let cg = self.glyph_cache.entry((pg.cache_key, gc)).or_insert_with(|| {
                    if let Some(img) = swash_cache.get_image(font_system, pg.cache_key) {
                        let mut p = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)).unwrap();
                        let (r, g, b, a) = (gc.r(), gc.g(), gc.b(), gc.a());
                        for (idx, &alpha) in img.data.iter().enumerate() {
                            let af = (alpha as f32 / 255.0) * (a as f32 / 255.0);
                            p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                        }
                        CachedGlyph { pixmap: p, left: img.placement.left, top: img.placement.top }
                    } else { CachedGlyph { pixmap: Pixmap::new(1, 1).unwrap(), left: 0, top: 0 } }
                });
                pixmap.draw_pixmap(pg.x + cg.left, pg.y - cg.top, cg.pixmap.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                
                if is_partner {
                    let mut bp = Paint::default(); bp.set_color_rgba8(0x56, 0x9c, 0xd6, 0xff);
                    pixmap.fill_rect(Rect::from_xywh(x_off + glyph.x, current_y_off + line_top + line_height - 2.0, glyph.w, 2.0).unwrap(), &bp, Transform::identity(), None);
                }
            }
        }

        if let Some((cx, cy)) = self.editor.cursor_position() {
            let mut cursor_render_info = None;
            self.editor.with_buffer(|buffer| {
                for run in buffer.layout_runs() {
                    if cy >= run.line_top as i32 && cy < (run.line_top + run.line_height) as i32 {
                        if !self.is_line_hidden(run.line_i) {
                            let y_shift = self.get_visual_y_shift(run.line_i);
                            cursor_render_info = Some(y_shift);
                        }
                        break;
                    }
                }
            });
            if let Some(y_shift) = cursor_render_info {
                let mut cp = Paint::default(); cp.set_color_rgba8(0xff, 0xff, 0xff, 0xff);
                pixmap.fill_rect(Rect::from_xywh(x_off + cx as f32, y_off - self.scroll_y - y_shift + cy as f32, 2.0, self.metrics.line_height).unwrap(), &cp, Transform::identity(), None);
            }
        }
    }

    pub fn handle_mouse(&mut self, font_system: &mut FontSystem, x: f32, y: f32, rect: Rect, click: Option<u32>) {
        let (x_off, y_off) = self.get_offsets(rect);
        
        let mut toggle_li = None;
        if let Some(1) = click {
            if x < x_off && x > rect.left() {
                self.editor.with_buffer(|buffer| {
                    for run in buffer.layout_runs() {
                        if self.is_line_hidden(run.line_i) { continue; }
                        let y_shift = self.get_visual_y_shift(run.line_i);
                        let vy = y_off - self.scroll_y - y_shift + run.line_top;
                        if y >= vy && y < vy + run.line_height {
                            toggle_li = Some(run.line_i);
                            break;
                        }
                    }
                });
            }
        }
        if let Some(li) = toggle_li { self.my_editor.toggle_fold(li); return; }

        let visual_y = y - y_off + self.scroll_y;
        let mut total_shift = 0.0;
        self.editor.with_buffer(|buffer| {
            for run in buffer.layout_runs() {
                if self.is_line_hidden(run.line_i) { continue; }
                let y_shift = self.get_visual_y_shift(run.line_i);
                let current_run_y = run.line_top - y_shift;
                if visual_y >= current_run_y && visual_y < (current_run_y + run.line_height) {
                    total_shift = y_shift;
                    break;
                }
            }
        });

        let ex = (x - x_off) as i32;
        let ey = (y - y_off + total_shift + self.scroll_y) as i32;
        
        if let Some(count) = click {
            match count {
                1 => self.editor.action(font_system, Action::Click { x: ex, y: ey }),
                2 => self.editor.action(font_system, Action::DoubleClick { x: ex, y: ey }),
                3 => self.editor.action(font_system, Action::TripleClick { x: ex, y: ey }),
                _ => {}
            }
        } else {
            self.editor.action(font_system, Action::Drag { x: ex, y: ey });
        }
    }

    pub fn sync(&mut self) {
        if self.needs_reshape {
            let text: String = self.editor.with_buffer(|b| b.lines.iter().map(|l| format!("{}\n", l.text())).collect());
            self.my_editor.rope = ropey::Rope::from_str(&text);
            self.my_editor.retokenize_all(&self.lang_def);
        }
    }
}

struct App {
    window: Option<Arc<Window>>,
    context: Option<Context<Arc<Window>>>,
    surface: Option<Surface<Arc<Window>, Arc<Window>>>,
    font_system: FontSystem,
    swash_cache: SwashCache,
    pixmap: Option<Pixmap>,
    editor_widget: Option<CodeEditorWidget>,
    all_languages: Vec<String>,
    current_lang: String,
    is_picker_open: bool,
    clipboard: Option<Clipboard>,
    modifiers: winit::event::Modifiers,
    last_click_time: Instant,
    click_count: u32,
    mouse_pos: (f32, f32),
    is_dragging: bool,
    needs_redraw: bool,
}

impl App {
    fn new(my_editor: MyEditor) -> Self {
        let mut all_languages = Vec::new();
        if let Ok(entries) = fs::read_dir("basic-languages") {
            for entry in entries.flatten() {
                if entry.file_type().map(|t| t.is_dir()).unwrap_or(false) {
                    if let Some(name) = entry.file_name().to_str() { all_languages.push(name.to_string()); }
                }
            }
        }
        all_languages.sort();
        Self {
            window: None, context: None, surface: None,
            font_system: FontSystem::new(), swash_cache: SwashCache::new(), pixmap: None,
            editor_widget: None, all_languages, current_lang: "rust".to_string(), is_picker_open: false,
            clipboard: Clipboard::new().ok(), modifiers: winit::event::Modifiers::default(),
            last_click_time: Instant::now(), click_count: 0, mouse_pos: (0.0, 0.0),
            is_dragging: false, needs_redraw: true,
        }
    }

    fn render(&mut self) {
        let (surface, pixmap) = match (&mut self.surface, &mut self.pixmap) {
            (Some(s), Some(p)) => (s, p),
            _ => return,
        };
        pixmap.fill(SkiaColor::from_rgba8(0x1e, 0x1e, 0x1e, 0xff));
        
        let editor_rect = Rect::from_xywh(0.0, UI_BAR_HEIGHT * SCALE, pixmap.width() as f32, pixmap.height() as f32 - UI_BAR_HEIGHT * SCALE).unwrap();
        if let Some(editor) = &mut self.editor_widget {
           editor.render(pixmap, &mut self.font_system, &mut self.swash_cache, editor_rect);
        }

        let mut bar_paint = Paint::default(); bar_paint.set_color_rgba8(0x2d, 0x2d, 0x2d, 0xff);
        pixmap.fill_rect(Rect::from_xywh(0.0, 0.0, pixmap.width() as f32, UI_BAR_HEIGHT * SCALE).unwrap(), &bar_paint, Transform::identity(), None);
        let label = format!("Language: {}", self.current_lang);
        App::draw_ui_text(pixmap, &mut self.font_system, &mut self.swash_cache, &label, 15.0 * SCALE, (UI_BAR_HEIGHT * 0.25) * SCALE, Color::rgb(0xee, 0xee, 0xee));

        if self.is_picker_open {
            App::render_picker_internal(pixmap, &mut self.font_system, &mut self.swash_cache, &self.all_languages, self.mouse_pos);
        }

        let mut buffer = surface.buffer_mut().unwrap();
        let pixels = pixmap.pixels();
        for i in 0..pixels.len() {
            let p = pixels[i];
            buffer[i] = (p.red() as u32) << 16 | (p.green() as u32) << 8 | (p.blue() as u32);
        }
        buffer.present().unwrap();
        self.needs_redraw = false;
    }

    fn render_picker_internal(pixmap: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, all_languages: &[String], mouse_pos: (f32, f32)) {
        let cols = 4;
        let p_w = cols as f32 * 150.0 * SCALE;
        let p_h = ((all_languages.len() as f32 / cols as f32).ceil() as usize) as f32 * 30.0 * SCALE;
        let mut bg = Paint::default(); bg.set_color_rgba8(0x33, 0x33, 0x33, 0xf8);
        pixmap.fill_rect(Rect::from_xywh(10.0, (UI_BAR_HEIGHT + 5.0) * SCALE, p_w, p_h).unwrap(), &bg, Transform::identity(), None);

        for (i, lang) in all_languages.iter().enumerate() {
            let lx = 10.0 + (i % cols) as f32 * 150.0 * SCALE + 10.0;
            let ly = (UI_BAR_HEIGHT + 5.0 + (i / cols) as f32 * 30.0) * SCALE + 5.0;
            let is_hover = mouse_pos.0 >= lx && mouse_pos.0 <= lx + 150.0 * SCALE && mouse_pos.1 >= ly && mouse_pos.1 <= ly + 30.0 * SCALE;
            let color = if is_hover { Color::rgb(0x56, 0x9c, 0xd6) } else { Color::rgb(0xbb, 0xbb, 0xbb) };
            App::draw_ui_text(pixmap, fs, sc, lang, lx, ly, color);
        }
    }

    fn draw_ui_text(pixmap: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, text: &str, x: f32, y: f32, color: Color) {
        let mut lab = Buffer::new(fs, Metrics::new(14.0, 20.0).scale(SCALE));
        lab.set_text(fs, text, &Attrs::new().family(Family::Monospace).color(color), Shaping::Advanced, None);
        lab.shape_until_scroll(fs, false);
        for run in lab.layout_runs() {
            for glyph in run.glyphs {
                let pg = glyph.physical((x, y + run.line_y), 1.0);
                if let Some(img) = sc.get_image(fs, pg.cache_key) {
                    let mut p = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)).unwrap();
                    let (r, g, b, a) = (color.r(), color.g(), color.b(), color.a());
                    for (idx, &alpha) in img.data.iter().enumerate() {
                        let af = (alpha as f32 / 255.0) * (a as f32 / 255.0);
                        p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                    }
                    pixmap.draw_pixmap(pg.x + img.placement.left, pg.y - img.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                }
            }
        }
    }
}

impl ApplicationHandler for App {
    fn resumed(&mut self, event_loop: &winit::event_loop::ActiveEventLoop) {
        let window = Arc::new(event_loop.create_window(WindowAttributes::default()
            .with_title("Vybe Editor (Professional)")
            .with_inner_size(winit::dpi::LogicalSize::new(1000.0, 800.0))).unwrap());
        let context = Context::new(window.clone()).unwrap();
        let surface = Surface::new(&context, window.clone()).unwrap();
        let size = window.inner_size();
        let lang = load_language("rust").unwrap();
        let my_editor = MyEditor::from_text("// Vybe Editor\nfn main() {\n    println!(\"Hello Vybe!\");\n}", &lang);
        self.editor_widget = Some(CodeEditorWidget::new(my_editor, &mut self.font_system));
        self.window = Some(window); self.context = Some(context); self.surface = Some(surface);
        self.pixmap = Some(Pixmap::new(size.width, size.height).unwrap());
        self.surface.as_mut().unwrap().resize(NonZeroU32::new(size.width).unwrap(), NonZeroU32::new(size.height).unwrap()).unwrap();
    }

    fn window_event(&mut self, event_loop: &winit::event_loop::ActiveEventLoop, _id: winit::window::WindowId, event: WindowEvent) {
        match event {
            WindowEvent::CloseRequested => event_loop.exit(),
            WindowEvent::ModifiersChanged(m) => self.modifiers = m,
            WindowEvent::Resized(size) => {
                if let (Some(surface), Some(window)) = (&mut self.surface, &self.window) {
                    if size.width > 0 && size.height > 0 {
                        surface.resize(NonZeroU32::new(size.width).unwrap(), NonZeroU32::new(size.height).unwrap()).unwrap();
                        self.pixmap = Some(Pixmap::new(size.width, size.height).unwrap());
                        window.request_redraw();
                    }
                }
            }
            WindowEvent::MouseWheel { delta, .. } => {
                let amount = match delta {
                    MouseScrollDelta::LineDelta(_, y) => y * 40.0,
                    MouseScrollDelta::PixelDelta(pos) => pos.y as f32,
                };
                if let Some(widget) = &mut self.editor_widget {
                    widget.scroll_y = (widget.scroll_y - amount).max(0.0);
                    self.needs_redraw = true; self.window.as_ref().unwrap().request_redraw();
                }
            }
            WindowEvent::KeyboardInput { event, .. } => {
                if event.state == ElementState::Pressed {
                    let mut acted = true;
                    let widget = self.editor_widget.as_mut().unwrap();
                    let cmd = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                    match event.key_without_modifiers() {
                        Key::Named(NamedKey::Backspace) => widget.editor.action(&mut self.font_system, Action::Backspace),
                        Key::Named(NamedKey::Delete) => widget.editor.action(&mut self.font_system, Action::Delete),
                        Key::Named(NamedKey::Enter) => widget.editor.action(&mut self.font_system, Action::Enter),
                        Key::Named(NamedKey::Tab) => widget.editor.action(&mut self.font_system, Action::Indent),
                        Key::Named(NamedKey::ArrowLeft) => widget.editor.action(&mut self.font_system, Action::Motion(Motion::Left)),
                        Key::Named(NamedKey::ArrowRight) => widget.editor.action(&mut self.font_system, Action::Motion(Motion::Right)),
                        Key::Named(NamedKey::ArrowUp) => widget.editor.action(&mut self.font_system, Action::Motion(Motion::Up)),
                        Key::Named(NamedKey::ArrowDown) => widget.editor.action(&mut self.font_system, Action::Motion(Motion::Down)),
                        Key::Character(c) if cmd && (c == "c" || c == "C") => { 
                            if let Some(t) = widget.editor.copy_selection() { if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } }
                        }
                        Key::Character(c) if cmd && (c == "v" || c == "V") => { 
                            if let Some(cb) = &mut self.clipboard { if let Ok(t) = cb.get_text() { for ch in t.chars() { widget.editor.action(&mut self.font_system, Action::Insert(ch)); } } }
                        }
                        Key::Character(c) if cmd && (c == "x" || c == "X") => {
                            if let Some(t) = widget.editor.copy_selection() {
                                if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); }
                                widget.editor.action(&mut self.font_system, Action::Delete);
                            }
                        }
                        Key::Character(c) if cmd && (c == "a" || c == "A") => {
                            widget.editor.action(&mut self.font_system, Action::Motion(Motion::BufferStart));
                            let mut last_y = 0.0;
                            widget.editor.with_buffer(|b| { if let Some(r) = b.layout_runs().last() { last_y = r.line_top + r.line_height; } });
                            widget.editor.action(&mut self.font_system, Action::Drag { x: 999999, y: last_y as i32 });
                        }
                        _ => {
                            if let Some(t) = event.text {
                                if !cmd { 
                                    for ch in t.chars() { 
                                        if !ch.is_control() { 
                                            widget.editor.action(&mut self.font_system, Action::Insert(ch)); 
                                            let closer = match ch {
                                                '(' => Some(')'), '{' => Some('}'), '[' => Some(']'), '"' => Some('"'), '\'' => Some('\''),
                                                _ => None
                                            };
                                            if let Some(c) = closer {
                                                widget.editor.action(&mut self.font_system, Action::Insert(c));
                                                widget.editor.action(&mut self.font_system, Action::Motion(Motion::Left));
                                            }
                                        } 
                                    } 
                                }
                                else { acted = false; }
                            } else { acted = false; }
                        }
                    }
                    if acted { widget.needs_reshape = true; widget.sync(); self.needs_redraw = true; self.window.as_ref().unwrap().request_redraw(); }
                }
            }
            WindowEvent::CursorMoved { position, .. } => {
                self.mouse_pos = (position.x as f32, position.y as f32);
                if self.is_dragging {
                    let rect = Rect::from_xywh(0.0, UI_BAR_HEIGHT * SCALE, self.pixmap.as_ref().unwrap().width() as f32, self.pixmap.as_ref().unwrap().height() as f32).unwrap();
                    self.editor_widget.as_mut().unwrap().handle_mouse(&mut self.font_system, self.mouse_pos.0, self.mouse_pos.1, rect, None);
                    self.needs_redraw = true; self.window.as_ref().unwrap().request_redraw();
                } else if self.is_picker_open { self.needs_redraw = true; self.window.as_ref().unwrap().request_redraw(); }
            }
            WindowEvent::MouseInput { state, button, .. } => {
                if button == MouseButton::Left && state == ElementState::Pressed {
                    if self.mouse_pos.1 < UI_BAR_HEIGHT * SCALE { self.is_picker_open = !self.is_picker_open; self.needs_redraw = true; self.window.as_ref().unwrap().request_redraw(); return; }
                    if self.is_picker_open {
                        for (i, lang) in self.all_languages.iter().enumerate() {
                            let lx = 10.0 + (i % 4) as f32 * 150.0 * SCALE + 10.0;
                            let ly = (UI_BAR_HEIGHT + 5.0 + (i / 4) as f32 * 30.0) * SCALE + 5.0;
                            if self.mouse_pos.0 >= lx && self.mouse_pos.0 <= lx + 150.0 * SCALE && self.mouse_pos.1 >= ly && self.mouse_pos.1 <= ly + 30.0 * SCALE {
                                self.current_lang = lang.clone(); self.editor_widget.as_mut().unwrap().set_language(lang);
                                self.is_picker_open = false; self.needs_redraw = true; self.window.as_ref().unwrap().request_redraw(); return;
                            }
                        }
                        self.is_picker_open = false; self.needs_redraw = true; self.window.as_ref().unwrap().request_redraw(); return;
                    }
                    let r = Rect::from_xywh(0.0, UI_BAR_HEIGHT * SCALE, self.pixmap.as_ref().unwrap().width() as f32, self.pixmap.as_ref().unwrap().height() as f32).unwrap();
                    self.click_count = if Instant::now().duration_since(self.last_click_time) < Duration::from_millis(500) { (self.click_count % 3) + 1 } else { 1 };
                    self.last_click_time = Instant::now();
                    self.editor_widget.as_mut().unwrap().handle_mouse(&mut self.font_system, self.mouse_pos.0, self.mouse_pos.1, r, Some(self.click_count));
                    self.is_dragging = true; self.needs_redraw = true; self.window.as_ref().unwrap().request_redraw();
                } else if button == MouseButton::Left && state == ElementState::Released { self.is_dragging = false; }
            }
            WindowEvent::RedrawRequested => self.render(),
            _ => {}
        }
    }
}

pub fn run_gui(my_editor: MyEditor) {
    let event_loop = EventLoop::new().unwrap();
    event_loop.set_control_flow(ControlFlow::Wait);
    let mut app = App::new(my_editor);
    event_loop.run_app(&mut app).unwrap();
}
