use ropey::Rope;
use crate::language::LanguageDef;

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
    pub collapsed_starts: std::collections::HashSet<usize>, // line indices that are COLLAPSED
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
        let mut ed = Self { rope, line_tokens: Vec::new(), line_states: Vec::new(), folds: Vec::new(), collapsed_starts: std::collections::HashSet::new() };
        ed.retokenize_all(lang);
        ed
    }

    // Expose rope for simple manipulations from renderer for now
    pub fn rope(&self) -> &Rope { &self.rope }

    /// Remove a single char at the given byte position (if valid).
    pub fn remove_char_at_byte(&mut self, byte_pos: usize, lang: &LanguageDef) {
        if byte_pos == 0 { return; }
        let char_pos = self.rope.byte_to_char(byte_pos);
        if char_pos == 0 { return; }
        // remove previous char
        let remove_from = char_pos - 1;
        self.rope.remove(remove_from..remove_from+1);
        // retokenize affected lines
        let line = self.rope.char_to_line(remove_from);
        let end_line = std::cmp::min(self.rope.len_lines() - 1, line + 3);
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
}
