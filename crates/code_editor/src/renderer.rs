use std::collections::HashMap;
use std::time::{Duration, Instant};
use std::num::NonZeroU32;
use std::sync::Arc;
use cosmic_text::{Attrs, Buffer, Color, Family, FontSystem, Metrics, Motion, Shaping, SwashCache, Action, Edit, AttrsList};
use tiny_skia::{Color as SkiaColor, Paint, Pixmap, PixmapPaint, Rect, Transform, ColorU8};
use winit::application::ApplicationHandler;
use winit::event::{WindowEvent, ElementState, MouseButton};
use winit::event_loop::{ControlFlow, EventLoop};
use winit::window::{Window, WindowAttributes};
use winit::keyboard::{Key, NamedKey};
#[cfg(target_os = "macos")]
use winit::platform::modifier_supplement::KeyEventExtModifierSupplement;
use softbuffer::{Context, Surface};
use arboard::Clipboard;

use crate::editor::{Editor as MyEditor, TokenKind};

const SIDEBAR_WIDTH: f32 = 70.0;
const SCALE: f32 = 2.0;
const TEXT_PADDING: f32 = 15.0;

struct CachedGlyph {
    pixmap: Pixmap,
    left: i32,
    top: i32,
}

pub fn apply_highlighting(editor: &mut cosmic_text::Editor<'static>, my_editor: &MyEditor, attrs: &Attrs) {
    let kw_color = Color::rgb(0x56, 0x9c, 0xd6);
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
                            if matches!(text.as_str(), "fn" | "let" | "mut" | "if" | "else" | "match" | "use" | "pub" | "struct" | "enum" | "trait" | "impl" | "return") {
                                kw_color
                            } else {
                                ident_color
                            }
                        }
                        _ => ident_color,
                    };
                    let start = token.start.saturating_sub(byte_offset);
                    let end = token.end.saturating_sub(byte_offset);
                    list.add_span(start..end, &attrs.clone().color(color));
                }
            }
            line.set_attrs_list(list);
            if li < my_editor.rope.len_lines() {
                byte_offset += my_editor.rope.line(li).len_bytes();
            } else {
                byte_offset += line.text().len() + 1;
            }
        }
    });
}

struct App {
    window: Option<Arc<Window>>,
    context: Option<Context<Arc<Window>>>,
    surface: Option<Surface<Arc<Window>, Arc<Window>>>,
    editor: Option<cosmic_text::Editor<'static>>,
    font_system: FontSystem,
    swash_cache: SwashCache,
    my_editor: MyEditor,
    metrics: Metrics,
    pixmap: Option<Pixmap>,
    glyph_cache: HashMap<(cosmic_text::CacheKey, Color), CachedGlyph>,
    digit_cache: Vec<CachedGlyph>,
    
    clipboard: Option<Clipboard>,
    modifiers: winit::event::Modifiers,
    last_click_time: Instant,
    click_count: u32,
    mouse_pos: (f32, f32),
    is_dragging: bool,
    needs_reshape: bool,
    needs_redraw: bool,
}

impl App {
    fn new(my_editor: MyEditor) -> Self {
        let mut font_system = FontSystem::new();
        let metrics = Metrics::new(14.0, 20.0).scale(SCALE);
        
        let digit_color = Color::rgb(0x85, 0x85, 0x85);
        let mut digit_cache = Vec::new();
        let mut swash_cache = SwashCache::new();

        for i in 0..10 {
            let mut lab = Buffer::new(&mut font_system, metrics);
            lab.set_text(&mut font_system, &format!("{}", i), &Attrs::new().family(Family::Monospace).color(digit_color), Shaping::Advanced, None);
            lab.shape_until_scroll(&mut font_system, false);
            if let Some(r) = lab.layout_runs().next() {
                if let Some(g) = r.glyphs.first() {
                    let pg = g.physical((0.0, 0.0), 1.0);
                    if let Some(img) = swash_cache.get_image(&mut font_system, pg.cache_key) {
                        let mut p = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)).unwrap();
                        let (r, g, b, a) = (digit_color.r(), digit_color.g(), digit_color.b(), digit_color.a());
                        for (idx, &alpha) in img.data.iter().enumerate() {
                            let af = (alpha as f32 / 255.0) * (a as f32 / 255.0);
                            p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                        }
                        digit_cache.push(CachedGlyph { pixmap: p, left: img.placement.left, top: img.placement.top });
                        continue;
                    }
                }
            }
            digit_cache.push(CachedGlyph { pixmap: Pixmap::new(1, 1).unwrap(), left: 0, top: 0 });
        }

        Self {
            window: None,
            context: None,
            surface: None,
            editor: None,
            font_system,
            swash_cache,
            my_editor,
            metrics,
            pixmap: None,
            glyph_cache: HashMap::new(),
            digit_cache,
            clipboard: Clipboard::new().ok(),
            modifiers: winit::event::Modifiers::default(),
            last_click_time: Instant::now(),
            click_count: 0,
            mouse_pos: (0.0, 0.0),
            is_dragging: false,
            needs_reshape: true,
            needs_redraw: true,
        }
    }

    fn render(&mut self) {
        let (_window, surface, editor, pixmap) = match (&self.window, &mut self.surface, &mut self.editor, &mut self.pixmap) {
            (Some(w), Some(s), Some(e), Some(p)) => (w, s, e, p),
            _ => return,
        };

        if self.needs_reshape {
            apply_highlighting(editor, &self.my_editor, &Attrs::new().family(Family::Monospace));
            editor.shape_as_needed(&mut self.font_system, false);
            self.needs_reshape = false;
        }

        pixmap.fill(SkiaColor::from_rgba8(0x1e, 0x1e, 0x1e, 0xff));
        let mut paint = Paint::default();
        paint.set_color_rgba8(0x2d, 0x2d, 0x2d, 0xff);
        pixmap.fill_rect(Rect::from_xywh(0.0, 0.0, SIDEBAR_WIDTH * SCALE, pixmap.height() as f32).unwrap(), &paint, Transform::identity(), None);

        let selection = editor.selection_bounds();
        let x_off = (SIDEBAR_WIDTH + TEXT_PADDING) * SCALE;

        editor.with_buffer(|buffer| {
            let mut last_para = None;
            for run in buffer.layout_runs() {
                if last_para != Some(run.line_i) {
                    let s = format!("{}", run.line_i + 1);
                    let mut digit_x = (SIDEBAR_WIDTH * SCALE) as i32 - 15;
                    for ch in s.chars().rev() {
                        if let Some(digit) = ch.to_digit(10) {
                            let cg = &self.digit_cache[digit as usize];
                            let y = run.line_top as i32 + (4.0 * SCALE) as i32;
                            pixmap.draw_pixmap(digit_x - cg.pixmap.width() as i32 + cg.left, y, cg.pixmap.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                            digit_x -= 10 * SCALE as i32;
                        }
                    }
                    last_para = Some(run.line_i);
                }

                if let Some((s_start, s_end)) = selection {
                    if let Some((hx, hw)) = run.highlight(s_start, s_end) {
                        let mut sp = Paint::default(); sp.set_color_rgba8(0x26, 0x4f, 0x78, 0xff);
                        pixmap.fill_rect(Rect::from_xywh(x_off + hx, run.line_top, hw, run.line_height).unwrap(), &sp, Transform::identity(), None);
                    }
                }

                for glyph in run.glyphs {
                    let pg = glyph.physical((x_off, 0.0), 1.0);
                    let gc = glyph.color_opt.unwrap_or(Color::rgb(0xee, 0xee, 0xee));
                    let swash = &mut self.swash_cache;
                    let fs = &mut self.font_system;
                    let cg = self.glyph_cache.entry((pg.cache_key, gc)).or_insert_with(|| {
                        if let Some(img) = swash.get_image(fs, pg.cache_key) {
                            let mut p = Pixmap::new(img.placement.width.max(1), img.placement.height.max(1)).unwrap();
                            let (r, g, b, a) = (gc.r(), gc.g(), gc.b(), gc.a());
                            for (idx, &alpha) in img.data.iter().enumerate() {
                                let af = (alpha as f32 / 255.0) * (a as f32 / 255.0);
                                p.pixels_mut()[idx] = ColorU8::from_rgba((r as f32 * af) as u8, (g as f32 * af) as u8, (b as f32 * af) as u8, (255.0 * af) as u8).premultiply();
                            }
                            CachedGlyph { pixmap: p, left: img.placement.left, top: img.placement.top }
                        } else { CachedGlyph { pixmap: Pixmap::new(1, 1).unwrap(), left: 0, top: 0 } }
                    });
                    pixmap.draw_pixmap(pg.x + cg.left, run.line_y as i32 + pg.y - cg.top, cg.pixmap.as_ref(), &PixmapPaint::default(), Transform::identity(), None);
                }
            }
        });

        if let Some((cx, cy)) = editor.cursor_position() {
            let mut cp = Paint::default(); cp.set_color_rgba8(0xff, 0xff, 0xff, 0xff);
            pixmap.fill_rect(Rect::from_xywh(x_off + cx as f32, cy as f32, 2.0, self.metrics.line_height).unwrap(), &cp, Transform::identity(), None);
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
}

impl ApplicationHandler for App {
    fn resumed(&mut self, event_loop: &winit::event_loop::ActiveEventLoop) {
        let window = Arc::new(event_loop.create_window(WindowAttributes::default()
            .with_title("Vybe Editor (Professional Clipboard Edition)")
            .with_inner_size(winit::dpi::LogicalSize::new(900.0, 700.0))).unwrap());
        
        let context = Context::new(window.clone()).unwrap();
        let surface = Surface::new(&context, window.clone()).unwrap();
        
        let mut buffer = Buffer::new(&mut self.font_system, self.metrics);
        buffer.set_size(&mut self.font_system, Some(900.0 * SCALE), Some(700.0 * SCALE));
        buffer.set_text(&mut self.font_system, &self.my_editor.rope.to_string(), &Attrs::new().family(Family::Monospace), Shaping::Advanced, None);
        
        self.editor = Some(cosmic_text::Editor::new(buffer));
        self.window = Some(window);
        self.context = Some(context);
        self.surface = Some(surface);
        
        let size = self.window.as_ref().unwrap().inner_size();
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
                        if let Some(editor) = &mut self.editor {
                            editor.with_buffer_mut(|b| b.set_size(&mut self.font_system, Some(size.width as f32), Some(size.height as f32)));
                        }
                        self.needs_redraw = true;
                        window.request_redraw();
                    }
                }
            }
            WindowEvent::KeyboardInput { event, .. } => {
                if event.state == ElementState::Pressed {
                    let mut acted = true;
                    let editor = self.editor.as_mut().unwrap();
                    let is_cmd = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                    
                    match event.key_without_modifiers() {
                        Key::Named(NamedKey::Backspace) => { editor.action(&mut self.font_system, Action::Backspace); self.needs_reshape = true; }
                        Key::Named(NamedKey::Delete) => { editor.action(&mut self.font_system, Action::Delete); self.needs_reshape = true; }
                        Key::Named(NamedKey::Enter) => { editor.action(&mut self.font_system, Action::Enter); self.needs_reshape = true; }
                        Key::Named(NamedKey::Tab) => { editor.action(&mut self.font_system, Action::Indent); self.needs_reshape = true; }
                        Key::Named(NamedKey::ArrowLeft) => { editor.action(&mut self.font_system, Action::Motion(Motion::Left)); }
                        Key::Named(NamedKey::ArrowRight) => { editor.action(&mut self.font_system, Action::Motion(Motion::Right)); }
                        Key::Named(NamedKey::ArrowUp) => { editor.action(&mut self.font_system, Action::Motion(Motion::Up)); }
                        Key::Named(NamedKey::ArrowDown) => { editor.action(&mut self.font_system, Action::Motion(Motion::Down)); }
                        Key::Character(c) if is_cmd && (c == "c" || c == "C") => {
                            if let Some(text) = editor.copy_selection() {
                                if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(text); }
                            }
                        }
                        Key::Character(c) if is_cmd && (c == "v" || c == "V") => {
                            if let Some(cb) = &mut self.clipboard {
                                if let Ok(text) = cb.get_text() {
                                    for ch in text.chars() { editor.action(&mut self.font_system, Action::Insert(ch)); }
                                    self.needs_reshape = true;
                                }
                            }
                        }
                        Key::Character(c) if is_cmd && (c == "x" || c == "X") => {
                            if let Some(text) = editor.copy_selection() {
                                if let Some(cb) = &mut self.clipboard { let _ = cb.set_text(text); }
                                editor.action(&mut self.font_system, Action::Delete);
                                self.needs_reshape = true;
                            }
                        }
                        Key::Character(c) if is_cmd && (c == "a" || c == "A") => {
                            editor.action(&mut self.font_system, Action::Motion(Motion::BufferStart));
                            editor.action(&mut self.font_system, Action::Drag { x: 999999, y: 999999 }); // approximation for end
                        }
                        _ => {
                            if let Some(text) = event.text {
                                if !is_cmd {
                                    for c in text.chars() {
                                        if !c.is_control() {
                                            editor.action(&mut self.font_system, Action::Insert(c));
                                            self.needs_reshape = true;
                                        }
                                    }
                                } else { acted = false; }
                            } else { acted = false; }
                        }
                    }
                    if acted { 
                        if self.needs_reshape {
                            let text: String = editor.with_buffer(|b| b.lines.iter().map(|l| format!("{}\n", l.text())).collect());
                            self.my_editor.rope = ropey::Rope::from_str(&text);
                            self.my_editor.retokenize_all();
                        }
                        self.needs_redraw = true; 
                        self.window.as_ref().unwrap().request_redraw();
                    }
                }
            }
            WindowEvent::CursorMoved { position, .. } => {
                self.mouse_pos = (position.x as f32, position.y as f32);
                if self.is_dragging {
                    let x_off = (SIDEBAR_WIDTH + TEXT_PADDING) * SCALE;
                    let (ex, ey) = ((self.mouse_pos.0 as f32 - x_off) as i32, self.mouse_pos.1 as i32);
                    self.editor.as_mut().unwrap().action(&mut self.font_system, Action::Drag { x: ex, y: ey });
                    self.needs_redraw = true;
                    self.window.as_ref().unwrap().request_redraw();
                }
            }
            WindowEvent::MouseInput { state, button, .. } => {
                if button == MouseButton::Left {
                    let x_off = (SIDEBAR_WIDTH + TEXT_PADDING) * SCALE;
                    let (ex, ey) = ((self.mouse_pos.0 as f32 - x_off) as i32, self.mouse_pos.1 as i32);
                    let editor = self.editor.as_mut().unwrap();
                    
                    if state == ElementState::Pressed {
                        let now = Instant::now();
                        self.click_count = if now.duration_since(self.last_click_time) < Duration::from_millis(500) { (self.click_count % 3) + 1 } else { 1 };
                        self.last_click_time = now;
                        match self.click_count {
                            1 => editor.action(&mut self.font_system, Action::Click { x: ex, y: ey }),
                            2 => editor.action(&mut self.font_system, Action::DoubleClick { x: ex, y: ey }),
                            3 => editor.action(&mut self.font_system, Action::TripleClick { x: ex, y: ey }),
                            _ => {}
                        }
                        self.is_dragging = true;
                    } else {
                        self.is_dragging = false;
                    }
                    self.needs_redraw = true;
                    self.window.as_ref().unwrap().request_redraw();
                }
            }
            WindowEvent::RedrawRequested => {
                self.render();
            }
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
