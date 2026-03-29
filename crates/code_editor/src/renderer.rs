use std::collections::{HashMap, HashSet};
use std::time::{Duration, Instant};
use std::num::NonZeroU32;
use std::sync::Arc;
use std::fs;
use cosmic_text::{Attrs, Buffer, Color, Family, FontSystem, Metrics, Motion, Shaping, SwashCache, Action, Edit, AttrsList, Cursor, Selection};
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
const MINIMAP_WIDTH: f32 = 80.0;
const UI_BAR_HEIGHT: f32 = 40.0;

#[derive(Clone, Copy)]
pub struct Theme {
    pub bg: Color,
    pub sidebar_bg: Color,
    pub current_line: Color,
    pub selection: Color,
    pub match_highlight: Color,
    pub text: Color,
    pub kw: Color,
    pub type_kw: Color,
    pub comment: Color,
    pub string: Color,
    pub number: Color,
    pub guide: Color,
    pub bracket: Color,
}

impl Theme {
    pub fn dark() -> Self {
        Self {
            bg: Color::rgb(0x1e, 0x1e, 0x1e),
            sidebar_bg: Color::rgb(0x2d, 0x2d, 0x2d),
            current_line: Color::rgb(0x32, 0x32, 0x32),
            selection: Color::rgb(0x26, 0x4f, 0x78),
            match_highlight: Color::rgb(0x3e, 0x44, 0x51),
            text: Color::rgb(0xee, 0xee, 0xee),
            kw: Color::rgb(0x56, 0x9c, 0xd6),
            type_kw: Color::rgb(0x4e, 0xc9, 0xb0),
            comment: Color::rgb(0x6a, 0x99, 0x55),
            string: Color::rgb(0xce, 0x91, 0x78),
            number: Color::rgb(0xb5, 0xce, 0xa8),
            guide: Color::rgb(0x40, 0x40, 0x40),
            bracket: Color::rgb(0x56, 0x9c, 0xd6),
        }
    }
}

struct CachedGlyph {
    pixmap: Pixmap,
    left: i32,
    top: i32,
}

pub fn apply_highlighting(editor: &mut cosmic_text::Editor<'static>, my_editor: &MyEditor, attrs: &Attrs, lang: &LanguageDef, theme: &Theme) {
    editor.with_buffer_mut(|buffer| {
        let mut byte_offset = 0usize;
        for (li, line) in buffer.lines.iter_mut().enumerate() {
            let mut list = AttrsList::new(attrs);
            if li < my_editor.line_tokens.len() {
                for token in &my_editor.line_tokens[li] {
                    let color = match token.kind {
                        TokenKind::LineComment | TokenKind::BlockComment => theme.comment,
                        TokenKind::String => theme.string,
                        TokenKind::Number => theme.number,
                        TokenKind::Punct => Color::rgb(0xd4, 0xd4, 0xd4),
                        TokenKind::Identifier => {
                            let text = my_editor.slice(token.start, token.end);
                            if lang.keywords.contains(&text) { theme.kw }
                            else if lang.type_keywords.contains(&text) { theme.type_kw }
                            else if lang.constants.contains(&text) { theme.number }
                            else { theme.text }
                        }
                        _ => theme.text,
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
    theme: Theme,
    metrics: Metrics,
    glyph_cache: HashMap<(cosmic_text::CacheKey, Color), CachedGlyph>,
    digit_cache: Vec<CachedGlyph>,
    needs_reshape: bool,
    pub scroll_y: f32,
    search_query: String,
    replace_query: String,
    is_search_open: bool,
    is_replace_open: bool,
    context_menu: Option<((f32, f32), Vec<String>)>,
}

impl CodeEditorWidget {
    pub fn new(my_editor: MyEditor, font_system: &mut FontSystem) -> Self {
        let metrics = Metrics::new(14.0, 20.0).scale(SCALE);
        let lang_def = load_language("rust").unwrap_or_else(|| LanguageDef {
            keywords: HashSet::new(), type_keywords: HashSet::new(), constants: HashSet::new(), operators: HashSet::new(), comments: None, brackets: Vec::new(),
        });
        let theme = Theme::dark();
        let mut editor_internal = my_editor;
        editor_internal.retokenize_all(&lang_def);
        let mut buffer = Buffer::new(font_system, metrics);
        buffer.set_text(font_system, &editor_internal.rope.to_string(), &Attrs::new().family(Family::Monospace), Shaping::Advanced, None);
        
        let mut swash_cache = SwashCache::new();
        let mut digit_cache = Vec::new();
        let digit_color = Color::rgb(0x85, 0x85, 0x85);
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
                        for (idx, &al) in img.data.iter().enumerate() { let af = (al as f32 / 255.0) * (a as f32 / 255.0); p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply(); }
                        digit_cache.push(CachedGlyph { pixmap: p, left: img.placement.left, top: img.placement.top }); break;
                    }
                }
            }
            if digit_cache.len() <= i { digit_cache.push(CachedGlyph { pixmap: Pixmap::new(1, 1).unwrap(), left: 0, top: 0 }); }
        }

        Self { editor: cosmic_text::Editor::new(buffer), my_editor: editor_internal, lang_def, theme, metrics, glyph_cache: HashMap::new(), digit_cache, needs_reshape: true, scroll_y: 0.0, search_query: String::new(), replace_query: String::new(), is_search_open: false, is_replace_open: false, context_menu: None }
    }

    pub fn set_language(&mut self, lang_name: &str) {
        if let Some(lang) = load_language(lang_name) {
            self.lang_def = lang; self.my_editor.retokenize_all(&self.lang_def); self.needs_reshape = true;
        }
    }

    fn get_offsets(&self, rect: Rect) -> (f32, f32) { (rect.left() + SIDEBAR_WIDTH * SCALE, rect.top()) }

    fn is_line_hidden(&self, li: usize) -> bool {
        for (s, e) in &self.my_editor.folds { if self.my_editor.collapsed_starts.contains(s) && li > *s && li <= *e { return true; } } false
    }

    fn get_visual_y_shift(&self, li: usize) -> f32 {
        let mut h = 0; for (s, e) in &self.my_editor.folds { if self.my_editor.collapsed_starts.contains(s) && li > *e { h += e - s; } }
        h as f32 * self.metrics.line_height
    }

    fn find_next(&mut self, _fs: &mut FontSystem) {
        if self.search_query.is_empty() { return; }
        let cursor = self.editor.cursor();
        let total = self.my_editor.rope.to_string();
        let mut start = 0;
        self.editor.with_buffer(|b| { for l in b.lines.iter().take(cursor.line) { start += l.text().len() + 1; } start += cursor.index; });
        let match_idx = total[start.min(total.len())..].find(&self.search_query).map(|i| i + start).or_else(|| total.find(&self.search_query));
        if let Some(idx) = match_idx {
            let mut cb = 0; let mut tl = 0; let mut tc = 0;
            for (li, text) in total.split('\n').enumerate() { if idx >= cb && idx <= cb + text.len() { tl = li; tc = idx - cb; break; } cb += text.len() + 1; }
            self.editor.set_cursor(Cursor::new(tl, tc)); self.editor.set_selection(Selection::Normal(Cursor::new(tl, tc + self.search_query.len()))); self.needs_reshape = true;
        }
    }

    pub fn replace_next(&mut self, fs: &mut FontSystem) {
        if self.search_query.is_empty() { return; }
        if let Some(sel) = self.editor.copy_selection() { if sel == self.search_query { for ch in self.replace_query.clone().chars() { self.editor.action(fs, Action::Insert(ch)); } self.find_next(fs); return; } }
        self.find_next(fs);
    }

    pub fn replace_all(&mut self, fs: &mut FontSystem) {
        if self.search_query.is_empty() { return; }
        let mut text = self.my_editor.rope.to_string();
        text = text.replace(&self.search_query, &self.replace_query);
        self.editor.with_buffer_mut(|b| b.set_text(fs, &text, &Attrs::new().family(Family::Monospace), Shaping::Advanced, None));
        self.my_editor.rope = ropey::Rope::from_str(&text); self.my_editor.retokenize_all(&self.lang_def); self.needs_reshape = true;
    }

    pub fn render(&mut self, pixmap: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, rect: Rect) {
        if self.needs_reshape { apply_highlighting(&mut self.editor, &self.my_editor, &Attrs::new().family(Family::Monospace), &self.lang_def, &self.theme); self.editor.shape_as_needed(fs, false); self.needs_reshape = false; }
        let (x_off, y_off) = self.get_offsets(rect);
        let mut sp = Paint::default(); sp.set_color_rgba8(self.theme.sidebar_bg.r(), self.theme.sidebar_bg.g(), self.theme.sidebar_bg.b(), self.theme.sidebar_bg.a());
        pixmap.fill_rect(Rect::from_xywh(rect.left(), rect.top(), SIDEBAR_WIDTH * SCALE, rect.height()).unwrap(), &sp, Transform::identity(), None);
        let cursor_state = self.editor.cursor(); let selection = self.editor.selection_bounds();
        let partner = self.my_editor.find_matching_bracket(cursor_state.line, cursor_state.index, &self.lang_def).or_else(|| if cursor_state.index > 0 { self.my_editor.find_matching_bracket(cursor_state.line, cursor_state.index - 1, &self.lang_def) } else { None });
        let mut total_h = 0.0; self.editor.with_buffer(|b| { for r in b.layout_runs() { if !self.is_line_hidden(r.line_i) { total_h += r.line_height; } } });
        self.scroll_y = self.scroll_y.clamp(0.0, (total_h - rect.height() + 100.0).max(0.0));
        if let Some((_, cy)) = self.editor.cursor_position() {
             let mut cli = 0; self.editor.with_buffer(|b| { for r in b.layout_runs() { if cy >= r.line_top as i32 && cy < (r.line_top + r.line_height) as i32 { cli = r.line_i; break; } } });
             let cys = cy as f32 - self.get_visual_y_shift(cli); let rcy = cys - self.scroll_y;
             if rcy < 0.0 { self.scroll_y = cys; } else if rcy > (rect.height() - self.metrics.line_height) { self.scroll_y = cys - (rect.height() - self.metrics.line_height); }
        }
        let mut runs = Vec::new();
        self.editor.with_buffer(|buffer| {
            for run in buffer.layout_runs() {
                if !self.is_line_hidden(run.line_i) {
                    let y_s = self.get_visual_y_shift(run.line_i); let v_t = run.line_top - y_s - self.scroll_y;
                    if v_t <= rect.height() && (v_t + run.line_height) >= 0.0 { runs.push((run.line_i, run.line_y, run.line_top, run.line_height, y_s, run.glyphs.to_vec(), run.text.to_string())); }
                }
            }
        });
        let mut cp = Paint::default(); cp.set_color_rgba8(self.theme.current_line.r(), self.theme.current_line.g(), self.theme.current_line.b(), self.theme.current_line.a());
        let mut last_para = None;
        for (i, ly, lt, lh, ys, glyphs, text) in runs {
            let cyo = y_off - ys - self.scroll_y;
            if i == cursor_state.line { pixmap.fill_rect(Rect::from_xywh(rect.left() + SIDEBAR_WIDTH * SCALE, cyo + lt, rect.width() - (SIDEBAR_WIDTH + MINIMAP_WIDTH) * SCALE, lh).unwrap(), &cp, Transform::identity(), None); }
            let mut gp = Paint::default(); gp.set_color_rgba8(self.theme.guide.r(), self.theme.guide.g(), self.theme.guide.b(), self.theme.guide.a());
            let tw = 4.0 * 8.4 * SCALE; let ls = text.chars().take_while(|c| c.is_whitespace()).count();
            for j in 1..=(ls/4) { pixmap.fill_rect(Rect::from_xywh(x_off + (j as f32 * tw), cyo + lt, 1.0, lh).unwrap(), &gp, Transform::identity(), None); }
            if last_para != Some(i) {
                let s = format!("{}", i + 1); let mut dx = (rect.left() + SIDEBAR_WIDTH * SCALE) as i32 - 15;
                for ch in s.chars().rev() { if let Some(d) = ch.to_digit(10) { let cg = &self.digit_cache[d as usize]; pixmap.draw_pixmap(dx - cg.pixmap.width() as i32 + cg.left, (cyo + ly) as i32 - cg.top, cg.pixmap.as_ref(), &PixmapPaint::default(), Transform::identity(), None); dx -= 10 * SCALE as i32; } }
                if self.my_editor.folds.iter().any(|(s, _)| *s == i) { let col = self.my_editor.collapsed_starts.contains(&i); App::draw_ui_text(pixmap, fs, sc, if col { "+" } else { "-" }, rect.left() + 5.0 * SCALE, cyo + ly - 14.0 * SCALE, if col { self.theme.kw } else { Color::rgb(0x85, 0x85, 0x85) }); }
                last_para = Some(i);
            }
            if let Some((ss, se)) = selection { self.editor.with_buffer(|b| { if let Some(r) = b.layout_runs().nth(i) { if let Some((hx, hw)) = r.highlight(ss, se) { let mut sp = Paint::default(); sp.set_color_rgba8(self.theme.selection.r(), self.theme.selection.g(), self.theme.selection.b(), self.theme.selection.a()); pixmap.fill_rect(Rect::from_xywh(x_off + hx, cyo + lt, hw, lh).unwrap(), &sp, Transform::identity(), None); } } }); }
            for g in &glyphs {
                let ip = partner.map(|(pl, pi)| i == pl && g.start == pi).unwrap_or(false);
                let pg = g.physical((x_off, cyo + ly), 1.0); let gc = g.color_opt.unwrap_or(self.theme.text);
                let cg = self.glyph_cache.entry((pg.cache_key, gc)).or_insert_with(|| {
                    if let Some(im) = sc.get_image(fs, pg.cache_key) {
                        let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap(); let (r, g, b, a) = (gc.r(), gc.g(), gc.b(), gc.a());
                        for (idx, &al) in im.data.iter().enumerate() { let af = (al as f32 / 255.0) * (a as f32 / 255.0); p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply(); }
                        CachedGlyph { pixmap: p, left: im.placement.left, top: im.placement.top }
                    } else { CachedGlyph { pixmap: Pixmap::new(1, 1).unwrap(), left: 0, top: 0 } }
                });
                pixmap.draw_pixmap(pg.x + cg.left, pg.y - cg.top, cg.pixmap.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                if ip { let mut bp = Paint::default(); bp.set_color_rgba8(self.theme.bracket.r(), self.theme.bracket.g(), self.theme.bracket.b(), self.theme.bracket.a()); pixmap.fill_rect(Rect::from_xywh(x_off + g.x, cyo + lt + lh - 2.0, g.w, 2.0).unwrap(), &bp, Transform::identity(), None); }
            }
        }
        let mx = rect.right() - MINIMAP_WIDTH * SCALE; let mut mp = Paint::default(); mp.set_color_rgba8(16, 16, 16, 255); pixmap.fill_rect(Rect::from_xywh(mx, rect.top(), MINIMAP_WIDTH * SCALE, rect.height()).unwrap(), &mp, Transform::identity(), None);
        let m_step = (rect.height() / self.my_editor.line_tokens.len().max(1) as f32).min(2.5); let mut m_y = rect.top();
        for li in 0..self.my_editor.line_tokens.len() {
            if self.is_line_hidden(li) { continue; }
            if let Some(tks) = self.my_editor.line_tokens.get(li) {
                let mut xp = mx + 2.0; for t in tks { let tc = match t.kind { TokenKind::Identifier => self.theme.kw, TokenKind::String => self.theme.string, TokenKind::LineComment | TokenKind::BlockComment => self.theme.comment, TokenKind::Number => self.theme.number, _ => self.theme.guide }; let mut tp = Paint::default(); tp.set_color_rgba8(tc.r(), tc.g(), tc.b(), 0xaa); let w = ((t.end - t.start) as f32 * 0.8).min(mx + MINIMAP_WIDTH * SCALE - xp); pixmap.fill_rect(Rect::from_xywh(xp, m_y, w, (m_step * 0.7).max(1.0)).unwrap(), &tp, Transform::identity(), None); xp += w + 1.0; if xp >= rect.right() { break; } }
            }
            m_y += m_step; if m_y > rect.bottom() { break; }
        }
        let v_h = (rect.height() / total_h.max(1.0)) * rect.height(); let v_y = (self.scroll_y / total_h.max(1.0)) * rect.height();
        let mut vp = Paint::default(); vp.set_color_rgba8(255, 255, 255, 17); pixmap.fill_rect(Rect::from_xywh(mx, rect.top() + v_y.min(rect.height() - v_h), MINIMAP_WIDTH * SCALE, v_h).unwrap(), &vp, Transform::identity(), None);
        if self.is_search_open {
             let mut sep = Paint::default(); sep.set_color_rgba8(51, 51, 51, 250);
             let eh = if self.is_replace_open { 60.0 } else { 30.0 } * SCALE;
             pixmap.fill_rect(Rect::from_xywh(rect.right() - 260.0 * SCALE, rect.top() + 10.0 * SCALE, 160.0 * SCALE, eh).unwrap(), &sep, Transform::identity(), None);
             App::draw_ui_text(pixmap, fs, sc, &format!("Find: {}", self.search_query), rect.right() - 250.0 * SCALE, rect.top() + 14.0 * SCALE, self.theme.text);
             if self.is_replace_open { App::draw_ui_text(pixmap, fs, sc, &format!("Replace: {}", self.replace_query), rect.right() - 250.0 * SCALE, rect.top() + 44.0 * SCALE, Color::rgb(206, 145, 120)); }
        }
        if let Some((pos, items)) = &self.context_menu {
            let mut mep = Paint::default(); mep.set_color_rgba8(45, 45, 45, 255);
            pixmap.fill_rect(Rect::from_xywh(pos.0, pos.1, 100.0 * SCALE, (items.len() as f32 * 25.0) * SCALE).unwrap(), &mep, Transform::identity(), None);
            for (i, item) in items.iter().enumerate() { App::draw_ui_text(pixmap, fs, sc, item, pos.0 + 10.0 * SCALE, pos.1 + (i as f32 * 25.0 + 5.0) * SCALE, self.theme.text); }
        }
        if let Some((cx, cy)) = self.editor.cursor_position() {
             let mut ci = None; self.editor.with_buffer(|b| { for r in b.layout_runs() { if cy >= r.line_top as i32 && cy < (r.line_top + r.line_height) as i32 && !self.is_line_hidden(r.line_i) { ci = Some(self.get_visual_y_shift(r.line_i)); } } });
             if let Some(ys) = ci { let mut cp = Paint::default(); cp.set_color_rgba8(255, 255, 255, 255); pixmap.fill_rect(Rect::from_xywh(x_off + cx as f32, y_off - self.scroll_y - ys + cy as f32, 2.0, self.metrics.line_height).unwrap(), &cp, Transform::identity(), None); }
        }
    }

    pub fn handle_mouse(&mut self, fs: &mut FontSystem, x: f32, y: f32, rect: Rect, click: Option<(u32, MouseButton, winit::event::Modifiers)>, clipboard: &mut Option<Clipboard>) {
        if let Some((_, MouseButton::Right, _)) = click { self.context_menu = Some(((x, y), vec!["Copy".to_string(), "Paste".to_string(), "Cut".to_string(), "Select All".to_string(), "Find".to_string(), "Replace".to_string()])); return; }
        if let Some((pos, _)) = self.context_menu {
            let menu_h = (6.0 * 25.0) * SCALE;
            if x >= pos.0 && x <= pos.0 + 100.0 * SCALE && y >= pos.1 && y <= pos.1 + menu_h {
                let idx = ((y - pos.1) / (25.0 * SCALE)) as usize; self.context_menu = None;
                match idx { 
                    0 => { if let Some(t) = self.editor.copy_selection() { if let Some(cb) = clipboard { let _ = cb.set_text(t); } } },
                    1 => { if let Some(cb) = clipboard { if let Ok(t) = cb.get_text() { for ch in t.chars() { self.editor.action(fs, Action::Insert(ch)); } self.needs_reshape = true; self.sync(); } } },
                    2 => { if let Some(t) = self.editor.copy_selection() { if let Some(cb) = clipboard { let _ = cb.set_text(t); } self.editor.action(fs, Action::Delete); self.needs_reshape = true; self.sync(); } },
                    3 => { self.editor.action(fs, Action::Motion(cosmic_text::Motion::BufferStart)); let mut ly = 0.0; self.editor.with_buffer(|b| if let Some(r) = b.layout_runs().last() { ly = r.line_top + r.line_height; }); self.editor.action(fs, Action::Drag { x: 999999, y: ly as i32 }); },
                    4 => { self.is_search_open = true; self.search_query.clear(); },
                    5 => { self.is_search_open = true; self.is_replace_open = true; self.search_query.clear(); self.replace_query.clear(); },
                    _ => {} 
                } return;
            } self.context_menu = None;
        }
        if x > rect.right() - MINIMAP_WIDTH * SCALE {
             let mry = (y - rect.top()) / rect.height(); let mut th = 0.0;
             self.editor.with_buffer(|b| { for r in b.layout_runs() { if !self.is_line_hidden(r.line_i) { th += r.line_height; } } });
             self.scroll_y = (mry * th).max(0.0); return;
        }
        let (x_off, y_off) = self.get_offsets(rect);
        if let Some((1, MouseButton::Left, _)) = click { if x < x_off && x > rect.left() { let mut fl = None; self.editor.with_buffer(|b| { for r in b.layout_runs() { if !self.is_line_hidden(r.line_i) { let ys = self.get_visual_y_shift(r.line_i); let vy = y_off - self.scroll_y - ys + r.line_top; if y >= vy && y < vy + r.line_height { fl = Some(r.line_i); break; } } } }); if let Some(li) = fl { self.my_editor.toggle_fold(li); return; } } }
        let vy = y - y_off + self.scroll_y; let mut ts = 0.0;
        self.editor.with_buffer(|b| { for r in b.layout_runs() { if !self.is_line_hidden(r.line_i) { let ys = self.get_visual_y_shift(r.line_i); if vy >= r.line_top - ys && vy < r.line_top - ys + r.line_height { ts = ys; break; } } } });
        let ex = (x - x_off) as i32; let ey = (y - y_off + ts + self.scroll_y) as i32;
        if let Some((count, _, mods)) = click {
            if !mods.state().shift_key() && count == 1 { self.editor.action(fs, Action::Escape); }
            match count { 1 => self.editor.action(fs, Action::Click { x: ex, y: ey }), 2 => self.editor.action(fs, Action::DoubleClick { x: ex, y: ey }), 3 => self.editor.action(fs, Action::TripleClick { x: ex, y: ey }), _ => {} }
        } else { self.editor.action(fs, Action::Drag { x: ex, y: ey }); }
    }

    pub fn sync(&mut self) { 
        if self.needs_reshape { 
            let mut t = String::new();
            self.editor.with_buffer(|b| {
                for (i, line) in b.lines.iter().enumerate() {
                    if i > 0 { t.push('\n'); }
                    t.push_str(line.text());
                }
            });
            self.my_editor.rope = ropey::Rope::from_str(&t); 
            self.my_editor.retokenize_all(&self.lang_def); 
        } 
    }
}

struct App {
    window: Option<Arc<Window>>, context: Option<Context<Arc<Window>>>, surface: Option<Surface<Arc<Window>, Arc<Window>>>, font_system: FontSystem, swash_cache: SwashCache, pixmap: Option<Pixmap>, editor_widget: Option<CodeEditorWidget>, all_languages: Vec<String>, current_lang: String, is_picker_open: bool, clipboard: Option<Clipboard>, modifiers: winit::event::Modifiers, last_click_time: Instant, click_count: u32, mouse_pos: (f32, f32), is_dragging: bool, needs_redraw: bool,
}

impl App {
    fn new(my_editor: MyEditor) -> Self {
        let mut langs = Vec::new(); if let Ok(es) = fs::read_dir("basic-languages") { for e in es.flatten() { if e.file_type().map(|t| t.is_dir()).unwrap_or(false) { if let Some(n) = e.file_name().to_str() { langs.push(n.to_string()); } } } } langs.sort();
        Self { window: None, context: None, surface: None, font_system: FontSystem::new(), swash_cache: SwashCache::new(), pixmap: None, editor_widget: None, all_languages: langs, current_lang: "rust".to_string(), is_picker_open: false, clipboard: Clipboard::new().ok(), modifiers: winit::event::Modifiers::default(), last_click_time: Instant::now(), click_count: 0, mouse_pos: (0.0, 0.0), is_dragging: false, needs_redraw: true }
    }
    fn render(&mut self) {
        let (surf, pix) = match (&mut self.surface, &mut self.pixmap) { (Some(s), Some(p)) => (s, p), _ => return }; pix.fill(SkiaColor::from_rgba8(30,30,30,255));
        let rect = Rect::from_xywh(0.0, UI_BAR_HEIGHT * SCALE, pix.width() as f32, pix.height() as f32 - UI_BAR_HEIGHT * SCALE).unwrap();
        if let Some(w) = &mut self.editor_widget { w.render(pix, &mut self.font_system, &mut self.swash_cache, rect); }
        let mut bp = Paint::default(); bp.set_color_rgba8(45,45,45,255); pix.fill_rect(Rect::from_xywh(0.0, 0.0, pix.width() as f32, UI_BAR_HEIGHT * SCALE).unwrap(), &bp, Transform::identity(), None);
        App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, &format!("Language: {}", self.current_lang), 15.0*SCALE, (UI_BAR_HEIGHT*0.25)*SCALE, Color::rgb(238,238,238));
        if self.is_picker_open { let cols = 4; let pw = cols as f32 * 150.0 * SCALE; let ph = ((self.all_languages.len() as f32 / cols as f32).ceil()) as f32 * 30.0 * SCALE; let mut bpg = Paint::default(); bpg.set_color_rgba8(51,51,51,248); pix.fill_rect(Rect::from_xywh(10.0, (UI_BAR_HEIGHT+5.0)*SCALE, pw, ph).unwrap(), &bpg, Transform::identity(), None); for (i, l) in self.all_languages.iter().enumerate() { let lx = 10.0 + (i%cols) as f32 * 150.0 * SCALE + 10.0; let ly = (UI_BAR_HEIGHT+5.0 + (i/cols) as f32 * 30.0) * SCALE + 5.0; let h = self.mouse_pos.0 >= lx && self.mouse_pos.0 <= lx + 150.0*SCALE && self.mouse_pos.1 >= ly && self.mouse_pos.1 <= ly + 30.0*SCALE; App::draw_ui_text(pix, &mut self.font_system, &mut self.swash_cache, l, lx, ly, if h { Color::rgb(86,156,214) } else { Color::rgb(187,187,187) }); } }
        let mut buffer = surf.buffer_mut().unwrap(); for (i, p) in pix.pixels().iter().enumerate() { buffer[i] = (p.red() as u32) << 16 | (p.green() as u32) << 8 | (p.blue() as u32); } buffer.present().unwrap(); self.needs_redraw = false;
    }
    fn draw_ui_text(pix: &mut Pixmap, fs: &mut FontSystem, sc: &mut SwashCache, text: &str, x: f32, y: f32, col: Color) {
        let mut lab = Buffer::new(fs, Metrics::new(14.0,20.0).scale(SCALE)); lab.set_text(fs, text, &Attrs::new().family(Family::Monospace).color(col), Shaping::Advanced, None); lab.shape_until_scroll(fs, false);
        for r in lab.layout_runs() { for g in r.glyphs { let pg = g.physical((x, y + r.line_y), 1.0); if let Some(im) = sc.get_image(fs, pg.cache_key) { let mut p = Pixmap::new(im.placement.width.max(1), im.placement.height.max(1)).unwrap(); let (cr, cg, cb, ca) = (col.r(), col.g(), col.b(), col.a()); for (idx, &al) in im.data.iter().enumerate() { let af = (al as f32 / 255.0) * (ca as f32 / 255.0); p.pixels_mut()[idx] = ColorU8::from_rgba((cr as f32 * af) as u8, (cg as f32 * af) as u8, (cb as f32 * af) as u8, (255.0 * af) as u8).premultiply(); } pix.draw_pixmap(pg.x + im.placement.left, pg.y - im.placement.top, p.as_ref(), &PixmapPaint::default(), Transform::identity(), None); } } }
    }
}

impl ApplicationHandler for App {
    fn resumed(&mut self, event_loop: &winit::event_loop::ActiveEventLoop) {
        let window = Arc::new(event_loop.create_window(WindowAttributes::default().with_title("Vybe IDE").with_inner_size(winit::dpi::LogicalSize::new(1100.0, 800.0))).unwrap());
        let ctx = Context::new(window.clone()).unwrap(); let surf = Surface::new(&ctx, window.clone()).unwrap(); let sz = window.inner_size();
        let lang = load_language("rust").expect("load rust");
        let my_editor = MyEditor::from_text("// Welcome to Vybe Professionals\nfn main() {\n    println!(\"Hello IDE!\");\n}", &lang);
        self.editor_widget = Some(CodeEditorWidget::new(my_editor, &mut self.font_system));
        self.window = Some(window); self.context = Some(ctx); self.surface = Some(surf);
        self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap()); self.surface.as_mut().unwrap().resize(NonZeroU32::new(sz.width).unwrap(), NonZeroU32::new(sz.height).unwrap()).unwrap();
    }
    fn window_event(&mut self, event_loop: &winit::event_loop::ActiveEventLoop, _id: winit::window::WindowId, event: WindowEvent) {
        match event {
            WindowEvent::CloseRequested => event_loop.exit(), WindowEvent::ModifiersChanged(m) => self.modifiers = m,
            WindowEvent::Resized(sz) => { if let (Some(s), Some(w)) = (&mut self.surface, &self.window) { if sz.width > 0 && sz.height > 0 { s.resize(NonZeroU32::new(sz.width).unwrap(), NonZeroU32::new(sz.height).unwrap()).expect("resize surface"); self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap()); w.request_redraw(); } } }
            WindowEvent::MouseWheel { delta, .. } => { let a = match delta { MouseScrollDelta::LineDelta(_, y) => y * 60.0, MouseScrollDelta::PixelDelta(pos) => pos.y as f32 }; if let Some(w) = &mut self.editor_widget { w.scroll_y -= a; self.window.as_ref().unwrap().request_redraw(); } }
            WindowEvent::KeyboardInput { event, .. } => {
                if event.state == ElementState::Pressed {
                    let mut acted = true; let w = self.editor_widget.as_mut().expect("has widget");
                    let cmd = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                    let alt = self.modifiers.state().alt_key(); let shift = self.modifiers.state().shift_key();
                    match event.key_without_modifiers() {
                        Key::Named(NamedKey::Backspace) => if w.is_search_open { if w.is_replace_open && alt { w.replace_query.pop(); } else { w.search_query.pop(); } } else { w.editor.action(&mut self.font_system, Action::Backspace); }
                        Key::Named(NamedKey::Enter) => if w.is_search_open { w.find_next(&mut self.font_system); } else { w.editor.action(&mut self.font_system, Action::Enter); }
                        Key::Named(NamedKey::Escape) => { w.is_search_open = false; w.context_menu = None; }
                        Key::Character(c) if cmd && (c == "f" || c == "F") => { w.is_search_open = true; w.search_query.clear(); }
                        Key::Character(c) if cmd && (c == "h" || c == "H") => { w.is_search_open = true; w.is_replace_open = !w.is_replace_open; }
                        Key::Character(c) if cmd && (c == "c" || c == "C") => { if let Some(t) = w.editor.copy_selection() { if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } } }
                        Key::Character(c) if cmd && (c == "v" || c == "V") => { if let Some(cb) = &mut self.clipboard { if let Ok(t) = cb.get_text() { for ch in t.chars() { w.editor.action(&mut self.font_system, Action::Insert(ch)); } } } }
                        Key::Character(c) if cmd && (c == "x" || c == "X") => { if let Some(t) = w.editor.copy_selection() { if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(t); } w.editor.action(&mut self.font_system, Action::Delete); } }
                        Key::Named(NamedKey::ArrowUp) if alt => { let li = w.editor.cursor().line; w.my_editor.move_line_up(li); w.needs_reshape = true; }
                        Key::Named(NamedKey::ArrowDown) if alt => { let li = w.editor.cursor().line; if shift { w.my_editor.duplicate_line(li); } else { w.my_editor.move_line_down(li); } w.needs_reshape = true; }
                        Key::Named(NamedKey::ArrowLeft) => w.editor.action(&mut self.font_system, Action::Motion(Motion::Left)), Key::Named(NamedKey::ArrowRight) => w.editor.action(&mut self.font_system, Action::Motion(Motion::Right)), Key::Named(NamedKey::ArrowUp) => w.editor.action(&mut self.font_system, Action::Motion(Motion::Up)), Key::Named(NamedKey::ArrowDown) => w.editor.action(&mut self.font_system, Action::Motion(Motion::Down)),
                        Key::Character(c) if cmd && (c == "a" || c == "A") => { w.editor.action(&mut self.font_system, Action::Motion(Motion::BufferStart)); let mut ly = 0.0; w.editor.with_buffer(|b| if let Some(r) = b.layout_runs().last() { ly = r.line_top + r.line_height; }); w.editor.action(&mut self.font_system, Action::Drag { x: 999999, y: ly as i32 }); }
                        _ => { if let Some(t) = event.text { if !cmd { for ch in t.chars() { if !ch.is_control() || ch == '\t' || ch == '\n' { if w.is_search_open { if w.is_replace_open && alt { w.replace_query.push(ch); } else { w.search_query.push(ch); } } else { w.editor.action(&mut self.font_system, Action::Insert(ch)); if let Some(cl) = match ch { '('=>Some(')'),'{'=>Some('}'),'['=>Some(']'),'"'=>Some('"'),'\''=>Some('\''),_=>None } { w.editor.action(&mut self.font_system, Action::Insert(cl)); w.editor.action(&mut self.font_system, Action::Motion(Motion::Left)); } } } } } else { acted = false; } } else { acted = false; } }
                    }
                    if acted { w.needs_reshape = true; w.sync(); self.window.as_ref().unwrap().request_redraw(); }
                }
            }
            WindowEvent::CursorMoved { position, .. } => { self.mouse_pos = (position.x as f32, position.y as f32); if self.is_dragging { let r = Rect::from_xywh(0.0, UI_BAR_HEIGHT * SCALE, self.pixmap.as_ref().unwrap().width() as f32, self.pixmap.as_ref().unwrap().height() as f32).unwrap(); self.editor_widget.as_mut().unwrap().handle_mouse(&mut self.font_system, self.mouse_pos.0, self.mouse_pos.1, r, None, &mut self.clipboard); self.window.as_ref().unwrap().request_redraw(); } }
            WindowEvent::MouseInput { state, button, .. } => {
                if state == ElementState::Pressed {
                    if self.mouse_pos.1 < UI_BAR_HEIGHT * SCALE { if button == MouseButton::Left { self.is_picker_open = !self.is_picker_open; self.window.as_ref().unwrap().request_redraw(); return; } }
                    if self.is_picker_open && button == MouseButton::Left { for (i, l) in self.all_languages.iter().enumerate() { let lx = 10.0 + (i % 4) as f32 * 150.0 * SCALE + 10.0; let ly = (UI_BAR_HEIGHT + 5.0 + (i / 4) as f32 * 30.0) * SCALE + 5.0; if self.mouse_pos.0 >= lx && self.mouse_pos.0 <= lx + 150.0 * SCALE && self.mouse_pos.1 >= ly && self.mouse_pos.1 <= ly + 30.0 * SCALE { self.current_lang = l.clone(); self.editor_widget.as_mut().unwrap().set_language(l); self.is_picker_open = false; self.window.as_ref().unwrap().request_redraw(); return; } } self.is_picker_open = false; self.window.as_ref().unwrap().request_redraw(); return; }
                    let r = Rect::from_xywh(0.0, UI_BAR_HEIGHT * SCALE, self.pixmap.as_ref().unwrap().width() as f32, self.pixmap.as_ref().unwrap().height() as f32).unwrap();
                    if button == MouseButton::Left { self.click_count = if Instant::now().duration_since(self.last_click_time) < Duration::from_millis(500) { (self.click_count % 3) + 1 } else { 1 }; self.last_click_time = Instant::now(); }
                    self.editor_widget.as_mut().unwrap().handle_mouse(&mut self.font_system, self.mouse_pos.0, self.mouse_pos.1, r, Some((self.click_count, button, self.modifiers)), &mut self.clipboard);
                    if button == MouseButton::Left { self.is_dragging = true; } self.window.as_ref().unwrap().request_redraw();
                } else { self.is_dragging = false; }
            }
            WindowEvent::RedrawRequested => self.render(),
            _ => {}
        }
    }
}

pub fn run_gui(my_editor: MyEditor) {
    let el = EventLoop::new().expect("event loop"); el.set_control_flow(ControlFlow::Wait);
    let mut app = App::new(my_editor); el.run_app(&mut app).expect("run app");
}
