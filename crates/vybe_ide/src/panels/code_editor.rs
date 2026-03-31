//! Code editor panel — basic text editing in the center area.

use cosmic_text::{Color as CosmicColor, FontSystem, SwashCache};
use tiny_skia::{Paint, Pixmap, Transform};

use crate::layout::Rect;
use crate::text::draw_mono;

pub struct CodeEditor {
    pub lines: Vec<String>,
    pub cursor_line: usize,
    pub cursor_col: usize,
    pub scroll_y: f32,
    pub line_height: f32,
    pub gutter_width: f32,
}

impl CodeEditor {
    pub fn new() -> Self {
        Self {
            lines: vec![String::new()],
            cursor_line: 0,
            cursor_col: 0,
            scroll_y: 0.0,
            line_height: 22.0,
            gutter_width: 55.0,
        }
    }

    pub fn set_code(&mut self, code: &str) {
        self.lines = code.lines().map(|l| l.to_string()).collect();
        if self.lines.is_empty() { self.lines.push(String::new()); }
        self.cursor_line = 0;
        self.cursor_col = 0;
        self.scroll_y = 0.0;
    }

    pub fn get_code(&self) -> String {
        self.lines.join("\n")
    }

    pub fn render(
        &self,
        pix: &mut Pixmap,
        fs: &mut FontSystem,
        sc: &mut SwashCache,
        rect: Rect,
        scale: f32,
    ) {
        let s = scale;
        let mut paint = Paint::default();

        // Background
        paint.set_color_rgba8(30, 30, 30, 255);
        fill(pix, &paint, rect.x, rect.y, rect.w, rect.h, s);

        let line_h = self.line_height;
        let gutter_w = self.gutter_width;
        let text_x = rect.x + gutter_w + 10.0;
        let visible_start = (self.scroll_y / line_h) as usize;
        let visible_count = (rect.h / line_h) as usize + 2;

        let line_num_color = CosmicColor::rgba(100, 100, 100, 255);
        let text_color = CosmicColor::rgba(212, 212, 212, 255);

        // Gutter background
        paint.set_color_rgba8(37, 37, 38, 255);
        fill(pix, &paint, rect.x, rect.y, gutter_w, rect.h, s);

        // Current line highlight
        let cur_y = rect.y + (self.cursor_line as f32 - visible_start as f32) * line_h;
        if cur_y >= rect.y && cur_y < rect.y + rect.h {
            paint.set_color_rgba8(40, 40, 40, 255);
            fill(pix, &paint, rect.x + gutter_w, cur_y, rect.w - gutter_w, line_h, s);
        }

        // Lines
        for i in visible_start..(visible_start + visible_count).min(self.lines.len()) {
            let y = rect.y + (i as f32 - visible_start as f32) * line_h;
            if y + line_h < rect.y || y > rect.y + rect.h { continue; }

            // Line number (right-aligned)
            let num_str = format!("{}", i + 1);
            let num_x = rect.x + gutter_w - 10.0 - num_str.len() as f32 * 8.0;
            draw_mono(pix, fs, sc, &num_str, num_x, y + 3.0, 13.0, line_num_color, s);

            // Code text
            draw_mono(pix, fs, sc, &self.lines[i], text_x, y + 3.0, 13.0, text_color, s);
        }

        // Cursor caret
        let char_w = 8.2; // monospace approximate
        let caret_x = text_x + self.cursor_col as f32 * char_w;
        let caret_y = rect.y + (self.cursor_line as f32 - visible_start as f32) * line_h;
        if caret_y >= rect.y && caret_y < rect.y + rect.h {
            paint.set_color_rgba8(220, 220, 220, 255);
            fill(pix, &paint, caret_x, caret_y, 1.5, line_h, s);
        }
    }

    pub fn handle_key(&mut self, key: &str) {
        match key {
            "Left" => {
                if self.cursor_col > 0 { self.cursor_col -= 1; }
                else if self.cursor_line > 0 {
                    self.cursor_line -= 1;
                    self.cursor_col = self.lines[self.cursor_line].len();
                }
            }
            "Right" => {
                let line_len = self.lines[self.cursor_line].len();
                if self.cursor_col < line_len { self.cursor_col += 1; }
                else if self.cursor_line + 1 < self.lines.len() {
                    self.cursor_line += 1;
                    self.cursor_col = 0;
                }
            }
            "Up" => {
                if self.cursor_line > 0 {
                    self.cursor_line -= 1;
                    self.cursor_col = self.cursor_col.min(self.lines[self.cursor_line].len());
                }
            }
            "Down" => {
                if self.cursor_line + 1 < self.lines.len() {
                    self.cursor_line += 1;
                    self.cursor_col = self.cursor_col.min(self.lines[self.cursor_line].len());
                }
            }
            "Home" => { self.cursor_col = 0; }
            "End" => { self.cursor_col = self.lines[self.cursor_line].len(); }
            "Enter" => {
                let line = self.lines[self.cursor_line].clone();
                let (before, after) = line.split_at(self.cursor_col.min(line.len()));
                self.lines[self.cursor_line] = before.to_string();
                self.lines.insert(self.cursor_line + 1, after.to_string());
                self.cursor_line += 1;
                self.cursor_col = 0;
            }
            "Backspace" => {
                if self.cursor_col > 0 {
                    let col = self.cursor_col - 1;
                    self.lines[self.cursor_line].remove(col);
                    self.cursor_col = col;
                } else if self.cursor_line > 0 {
                    let removed = self.lines.remove(self.cursor_line);
                    self.cursor_line -= 1;
                    self.cursor_col = self.lines[self.cursor_line].len();
                    self.lines[self.cursor_line].push_str(&removed);
                }
            }
            "Delete" => {
                let line_len = self.lines[self.cursor_line].len();
                if self.cursor_col < line_len {
                    self.lines[self.cursor_line].remove(self.cursor_col);
                } else if self.cursor_line + 1 < self.lines.len() {
                    let next = self.lines.remove(self.cursor_line + 1);
                    self.lines[self.cursor_line].push_str(&next);
                }
            }
            _ => {}
        }
    }

    pub fn insert_char(&mut self, ch: char) {
        let col = self.cursor_col.min(self.lines[self.cursor_line].len());
        self.lines[self.cursor_line].insert(col, ch);
        self.cursor_col = col + 1;
    }

    pub fn handle_click(&mut self, mx: f32, my: f32, rect: Rect) {
        if !rect.contains(mx, my) { return; }
        let visible_start = (self.scroll_y / self.line_height) as usize;
        let rel_y = my - rect.y;
        let line = visible_start + (rel_y / self.line_height) as usize;
        self.cursor_line = line.min(self.lines.len().saturating_sub(1));

        let text_x = rect.x + self.gutter_width + 10.0;
        let rel_x = mx - text_x;
        self.cursor_col = (rel_x / 8.2).max(0.0) as usize;
        self.cursor_col = self.cursor_col.min(self.lines[self.cursor_line].len());
    }

    pub fn scroll(&mut self, delta: f32, rect: Rect) {
        self.scroll_y = (self.scroll_y - delta * self.line_height * 3.0)
            .max(0.0)
            .min((self.lines.len() as f32 * self.line_height - rect.h).max(0.0));
    }
}

fn fill(pix: &mut Pixmap, paint: &Paint, x: f32, y: f32, w: f32, h: f32, s: f32) {
    if let Some(r) = tiny_skia::Rect::from_xywh(x * s, y * s, w * s, h * s) {
        pix.fill_rect(r, paint, Transform::identity(), None);
    }
}
