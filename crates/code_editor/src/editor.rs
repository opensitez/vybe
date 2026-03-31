use ropey::Rope;
use crate::language::LanguageDef;
use lsp_types::Diagnostic;
use std::collections::{HashSet, VecDeque};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LexerState {
    #[default]
    Normal,
    InBlockComment,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TokenKind {
    Whitespace,
    Identifier,
    Number,
    String,
    LineComment,
    BlockComment,
    Punct,
    #[allow(dead_code)]
    Unknown,
}

#[derive(Debug, Clone)]
pub struct TokenSpan {
    pub start: usize, // byte offsets
    pub end: usize,
    pub kind: TokenKind,
}

pub struct Editor {
    pub rope: Rope,
    pub line_tokens: Vec<Vec<TokenSpan>>,
    pub line_states: Vec<LexerState>,
    pub folds: Vec<(usize, usize)>, // foldable ranges: (start_li, end_li)
    pub collapsed_starts: HashSet<usize>, // line indices that are COLLAPSED
    pub diagnostics: Vec<Diagnostic>,
    pub history: VecDeque<(String, usize, usize)>,
    pub redo_history: VecDeque<(String, usize, usize)>,
}

// Tokenize a single line, returning tokens and the final state for the next line
fn tokenize_line(line: &str, base_offset: usize, lang: &LanguageDef, mut state: LexerState) -> (Vec<TokenSpan>, LexerState) {
    let bytes = line.as_bytes();
    let mut i = 0usize;
    let n = bytes.len();
    let mut out = Vec::new();

    let line_comment = lang.comments.as_ref().and_then(|c| c.line_comment.as_ref());
    let block_comment = lang.comments.as_ref().and_then(|c| c.block_comment.as_ref());
    
    while i < n {
        if state == LexerState::InBlockComment {
            let start = i;
            if let Some((_, end_marker)) = block_comment {
                let emb = end_marker.as_bytes();
                let mut found = false;
                while i < n {
                    if i + emb.len() <= n && &bytes[i..i+emb.len()] == emb {
                        i += emb.len();
                        state = LexerState::Normal;
                        found = true;
                        break;
                    }
                    i += 1;
                }
                out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::BlockComment });
                if found { continue; }
            } else { i = n; out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::BlockComment }); }
            continue;
        }

        let b = bytes[i];
        if b == b' ' || b == b'\t' || b == b'\r' || b == b'\n' {
            let start = i;
            i += 1;
            while i < n && (bytes[i] == b' ' || bytes[i] == b'\t' || bytes[i] == b'\r' || bytes[i] == b'\n') { i += 1; }
            out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::Whitespace });
            continue;
        }

        // Check Line Comment (Dynamic)
        if let Some(prefix) = line_comment {
            let pb = prefix.as_bytes();
            if i + pb.len() <= n && &bytes[i..i+pb.len()] == pb {
                let start = i;
                i = n;
                out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::LineComment });
                continue;
            }
        }

        // Check Block Comment Start (Dynamic)
        if let Some((start_marker, _)) = block_comment {
            let smb = start_marker.as_bytes();
            if i + smb.len() <= n && &bytes[i..i+smb.len()] == smb {
                state = LexerState::InBlockComment;
                // Re-loop will hit the InBlockComment case immediately
                continue;
            }
        }

        if b == b'\'' || b == b'"' {
            let quote = b;
            let start = i;
            i += 1;
            while i < n {
                if bytes[i] == b'\\' && i + 1 < n { i += 2; continue; }
                if bytes[i] == quote { i += 1; break; }
                i += 1;
            }
            out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::String });
            continue;
        }
        if b >= b'0' && b <= b'9' {
            let start = i;
            i += 1;
            while i < n && ((bytes[i] >= b'0' && bytes[i] <= b'9') || bytes[i] == b'.') { i += 1; }
            out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::Number });
            continue;
        }
        if (b >= b'a' && b <= b'z') || (b >= b'A' && b <= b'Z') || b == b'_' {
            let start = i;
            i += 1;
            while i < n && ((bytes[i] >= b'a' && bytes[i] <= b'z') || (bytes[i] >= b'A' && bytes[i] <= b'Z') || (bytes[i] >= b'0' && bytes[i] <= b'9') || bytes[i] == b'_') { i += 1; }
            out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::Identifier });
            continue;
        }
        let start = i;
        i += 1;
        out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::Punct });
    }
    (out, state)
}

// Redundant struct def removed

impl Editor {
    pub fn from_text(text: &str, lang: &LanguageDef) -> Self {
        let rope = Rope::from_str(text);
        let mut ed = Self { 
            rope, 
            line_tokens: Vec::new(), 
            line_states: Vec::new(), 
            folds: Vec::new(), 
            collapsed_starts: HashSet::new(), 
            diagnostics: Vec::new(),
            history: VecDeque::new(),
            redo_history: VecDeque::new(),
        };
        ed.retokenize_all(lang);
        ed
    }

    pub fn save_snapshot(&mut self, line: usize, col: usize) {
        let current = self.rope.to_string();
        // Avoid duplicate identical snapshots
        if let Some((last_text, _, _)) = self.history.back() {
            if last_text == &current { return; }
        }
        self.history.push_back((current, line, col));
        if self.history.len() > 100 { self.history.pop_front(); }
        self.redo_history.clear();
    }

    pub fn undo(&mut self, cur_line: usize, cur_col: usize) -> Option<(String, usize, usize)> {
        let current_text = self.rope.to_string();
        let (prev_text, prev_line, prev_col) = self.history.pop_back()?;
        
        self.redo_history.push_back((current_text, cur_line, cur_col));
        if self.redo_history.len() > 100 { self.redo_history.pop_front(); }
        
        self.rope = Rope::from_str(&prev_text);
        Some((prev_text, prev_line, prev_col))
    }

    pub fn redo(&mut self, cur_line: usize, cur_col: usize) -> Option<(String, usize, usize)> {
        let current_text = self.rope.to_string();
        let (next_text, next_line, next_col) = self.redo_history.pop_back()?;
        
        self.history.push_back((current_text, cur_line, cur_col));
        if self.history.len() > 100 { self.history.pop_front(); }
        
        self.rope = Rope::from_str(&next_text);
        Some((next_text, next_line, next_col))
    }

    // Expose rope for simple manipulations from renderer for now
    pub fn rope(&self) -> &Rope { &self.rope }

    /// Remove a range of text by byte positions.
    pub fn delete_range(&mut self, start_byte: usize, end_byte: usize, lang: &LanguageDef) {
        if start_byte >= end_byte { return; }
        let start_char = self.rope.byte_to_char(start_byte);
        let end_char = self.rope.byte_to_char(end_byte);
        self.rope.remove(start_char..end_char);
        
        // Retokenize affected area
        let line = self.rope.char_to_line(start_char);
        let end_line = std::cmp::min(self.rope.len_lines() - 1, line + 5);
        self.retokenize_range(line, end_line, lang);
    }

    pub fn slice(&self, start: usize, end: usize) -> String {
        let start_char = self.rope.byte_to_char(start);
        let end_char = self.rope.byte_to_char(end);
        self.rope.slice(start_char..end_char).to_string()
    }

    pub fn insert_str(&mut self, byte_pos: usize, s: &str, lang: &LanguageDef) {
        let char_pos = self.rope.byte_to_char(byte_pos);
        self.rope.insert(char_pos, s);
        let line = self.rope.char_to_line(char_pos);
        let end_line = std::cmp::min(self.rope.len_lines() - 1, line + 5);
        self.retokenize_range(line, end_line, lang);
    }

    pub fn insert_newline(&mut self, byte_pos: usize, lang: &LanguageDef) {
        let char_pos = self.rope.byte_to_char(byte_pos);
        let line_idx = self.rope.char_to_line(char_pos);
        let line = self.rope.line(line_idx).to_string();
        let indent = line.chars().take_while(|c| c.is_whitespace()).collect::<String>();
        
        // Smart Indent: maintain previous indentation
        let mut to_insert = String::from("\n");
        to_insert.push_str(&indent);
        
        // Block Indent: if previous line ended with an open bracket, add one more level
        let trimmed = line.trim_end();
        if trimmed.ends_with('{') || trimmed.ends_with('(') || trimmed.ends_with('[') || trimmed.ends_with(':') {
            to_insert.push_str("    ");
        }

        self.rope.insert(char_pos, &to_insert);
        let new_line = self.rope.char_to_line(char_pos + to_insert.len());
        self.retokenize_range(line_idx, new_line + 2, lang);
    }

    #[allow(dead_code)]
    pub fn replace_all(&mut self, from: &str, to: &str, lang: &LanguageDef) {
        if from.is_empty() { return; }
        let content = self.rope.to_string();
        let new_content = content.replace(from, to);
        if content != new_content {
            self.rope = Rope::from_str(&new_content);
            self.retokenize_all(lang);
        }
    }

    /// Retokenize the entire buffer and update line_tokens and folds.
    pub fn retokenize_all(&mut self, lang: &LanguageDef) {
        self.line_tokens.clear();
        self.line_states.clear();
        let mut byte_offset = 0usize;
        let mut state = LexerState::Normal;
        for line_idx in 0..self.rope.len_lines() {
            self.line_states.push(state);
            let line = self.rope.line(line_idx).to_string();
            let (tokens, next_state) = tokenize_line(&line, byte_offset, lang, state);
            self.line_tokens.push(tokens);
            state = next_state;
            byte_offset += line.len();
        }
        self.recompute_folds(lang);
    }

    /// Retokenize a range of lines (inclusive start..=end)
    pub fn retokenize_range(&mut self, start_line: usize, end_line: usize, lang: &LanguageDef) {
        if start_line >= self.rope.len_lines() { return; }
        
        let mut li = start_line;
        let mut byte_offset = 0usize;
        for i in 0..start_line { byte_offset += self.rope.line(i).len_bytes(); }
        
        let mut state = self.line_states[li];
        
        while li < self.rope.len_lines() {
            // Store the state at START of line
            if li < self.line_states.len() { self.line_states[li] = state; }
            else { self.line_states.push(state); }

            let line = self.rope.line(li).to_string();
            let (tokens, next_state) = tokenize_line(&line, byte_offset, lang, state);
            
            if li < self.line_tokens.len() { self.line_tokens[li] = tokens; }
            else { self.line_tokens.push(tokens); }

            byte_offset += line.len();
            li += 1;
            
            // If we've passed the requested range AND the state has stabilized, we can stop
            if li > end_line {
                if li < self.line_states.len() && self.line_states[li] == next_state { break; }
                if li == self.rope.len_lines() { break; }
            }
            state = next_state;
        }
        self.recompute_folds(lang);
    }

    /// Return a flat token list across all lines (useful for backends)
    pub fn tokenize_all(&self) -> Vec<TokenSpan> {
        let mut out = Vec::new();
        for line in &self.line_tokens {
            for t in line { out.push(t.clone()); }
        }
        out
    }

    pub fn toggle_fold(&mut self, line_idx: usize) {
        if self.collapsed_starts.contains(&line_idx) { 
            self.collapsed_starts.remove(&line_idx); 
        } else {
            if self.folds.iter().any(|(s, _)| *s == line_idx) {
                self.collapsed_starts.insert(line_idx);
            }
        }
    }

    fn recompute_folds(&mut self, lang: &LanguageDef) {
        self.folds.clear();
        if lang.brackets.is_empty() { return; }
        
        for (open, close) in &lang.brackets {
            let mut stack: Vec<usize> = Vec::new();
            for li in 0..self.rope.len_lines() {
                let line = self.rope.line(li).to_string();
                let trimmed = line.trim();
                if trimmed.contains(open) { stack.push(li); }
                if trimmed.contains(close) {
                    if let Some(start) = stack.pop() {
                        if li > start { self.folds.push((start, li)); }
                    }
                }
            }
        }
        self.folds.sort_by_key(|(s, _)| *s);
    }

    pub fn find_matching_bracket(&self, line_idx: usize, col: usize, lang: &LanguageDef) -> Option<(usize, usize)> {
        let line = self.rope.line(line_idx).to_string();
        if col >= line.len() { return None; }
        let ch = line.chars().nth(col)?;

        for (open, close) in &lang.brackets {
            let o_ch = open.chars().next()?;
            let c_ch = close.chars().next()?;
            if ch == o_ch {
                // Find closing
                let mut depth = 1;
                for li in line_idx..self.rope.len_lines() {
                    let text = self.rope.line(li).to_string();
                    let start_col = if li == line_idx { col + 1 } else { 0 };
                    for (ci, c) in text.chars().enumerate().skip(start_col) {
                        if c == o_ch { depth += 1; }
                        else if c == c_ch {
                            depth -= 1;
                            if depth == 0 { return Some((li, ci)); }
                        }
                    }
                }
            } else if ch == c_ch {
                // Find opening (backwards)
                let mut depth = 1;
                for li in (0..=line_idx).rev() {
                    let text = self.rope.line(li).to_string();
                    let chars: Vec<char> = text.chars().collect();
                    let end_col = if li == line_idx { col } else { chars.len() };
                    for ci in (0..end_col).rev() {
                        let c = chars[ci];
                        if c == c_ch { depth += 1; }
                        else if c == o_ch {
                            depth -= 1;
                            if depth == 0 { return Some((li, ci)); }
                        }
                    }
                }
            }
        }
        None
    }

    pub fn folds(&self) -> &Vec<(usize, usize)> { &self.folds }

    pub fn move_line_up(&mut self, line_idx: usize) {
        if line_idx == 0 || line_idx >= self.rope.len_lines() { return; }
        let current = self.rope.line(line_idx).to_string();
        let prev = self.rope.line(line_idx - 1).to_string();
        self.rope.remove(self.rope.line_to_char(line_idx - 1)..self.rope.line_to_char(line_idx + 1));
        self.rope.insert(self.rope.line_to_char(line_idx - 1), &format!("{}{}", current, prev));
    }

    pub fn move_line_down(&mut self, line_idx: usize) {
        if line_idx >= self.rope.len_lines() - 1 { return; }
        let current = self.rope.line(line_idx).to_string();
        let next = self.rope.line(line_idx + 1).to_string();
        self.rope.remove(self.rope.line_to_char(line_idx)..self.rope.line_to_char(line_idx + 2));
        self.rope.insert(self.rope.line_to_char(line_idx), &format!("{}{}", next, current));
    }

    pub fn duplicate_line(&mut self, line_idx: usize) {
        if line_idx >= self.rope.len_lines() { return; }
        let current = self.rope.line(line_idx).to_string();
        self.rope.insert(self.rope.line_to_char(line_idx + 1), &current);
    }

    pub fn insert_string(&mut self, byte_pos: usize, text: &str, lang: &LanguageDef) -> (usize, usize) {
        let char_pos = self.rope.byte_to_char(byte_pos);
        self.rope.insert(char_pos, text);
        
        // Handle line re-tokenization (roughly, for simplicity we re-tokenize the whole file for now 
        // to avoid sync issues, but normally we'd isolate it)
        self.retokenize_all(lang);
        
        let new_char_idx = char_pos + text.chars().count();
        let new_line = self.rope.char_to_line(new_char_idx);
        let new_col = new_char_idx - self.rope.line_to_char(new_line);
        (new_line, new_col)
    }
}
