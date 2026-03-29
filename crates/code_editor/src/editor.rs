use ropey::Rope;

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

// Tokenize a single line, returning tokens with absolute byte offsets (document-relative)
fn tokenize_line(line: &str, base_offset: usize) -> Vec<TokenSpan> {
    let bytes = line.as_bytes();
    let mut i = 0usize;
    let n = bytes.len();
    let mut out = Vec::new();
    while i < n {
        let b = bytes[i];
        if b == b' ' || b == b'\t' || b == b'\r' || b == b'\n' {
            let start = i;
            i += 1;
            while i < n && (bytes[i] == b' ' || bytes[i] == b'\t' || bytes[i] == b'\r' || bytes[i] == b'\n') { i += 1; }
            out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::Whitespace });
            continue;
        }
        if b == b'/' && i + 1 < n && bytes[i+1] == b'/' {
            let start = i;
            i += 2;
            while i < n && bytes[i] != b'\n' { i += 1; }
            out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::LineComment });
            continue;
        }
        if b == b'/' && i + 1 < n && bytes[i+1] == b'*' {
            let start = i;
            i += 2;
            while i + 1 < n && !(bytes[i] == b'*' && bytes[i+1] == b'/') { i += 1; }
            if i + 1 < n { i += 2; }
            out.push(TokenSpan { start: base_offset + start, end: base_offset + i, kind: TokenKind::BlockComment });
            continue;
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
    out
}

pub struct Editor {
    pub rope: Rope,
    // cached tokens per line (byte offsets relative to document start)
    pub line_tokens: Vec<Vec<TokenSpan>>,
    // simple fold ranges as (start_line, end_line)
    pub folds: Vec<(usize, usize)>,
}

impl Editor {
    pub fn from_text(text: &str) -> Self {
        let rope = Rope::from_str(text);
        let mut ed = Self { rope, line_tokens: Vec::new(), folds: Vec::new() };
        ed.retokenize_all();
        ed
    }

    // Expose rope for simple manipulations from renderer for now
    pub fn rope(&self) -> &Rope { &self.rope }

    /// Remove a single char at the given byte position (if valid).
    pub fn remove_char_at_byte(&mut self, byte_pos: usize) {
        if byte_pos == 0 { return; }
        let char_pos = self.rope.byte_to_char(byte_pos);
        if char_pos == 0 { return; }
        // remove previous char
        let remove_from = char_pos - 1;
        self.rope.remove(remove_from..remove_from+1);
        // retokenize affected lines
        let line = self.rope.char_to_line(remove_from);
        let end_line = std::cmp::min(self.rope.len_lines() - 1, line + 3);
        self.retokenize_range(line, end_line);
    }

    pub fn slice(&self, start: usize, end: usize) -> String {
        let start_char = self.rope.byte_to_char(start);
        let end_char = self.rope.byte_to_char(end);
        self.rope.slice(start_char..end_char).to_string()
    }

    pub fn insert_str(&mut self, byte_pos: usize, s: &str) {
        let char_pos = self.rope.byte_to_char(byte_pos);
        self.rope.insert(char_pos, s);
        // Retokenize affected lines: find line index for insertion
        let line = self.rope.char_to_line(char_pos);
        // retokenize this line and the next few lines to be safe
        let end_line = std::cmp::min(self.rope.len_lines() - 1, line + 5);
        self.retokenize_range(line, end_line);
    }

    /// Retokenize the entire buffer and update line_tokens and folds.
    pub fn retokenize_all(&mut self) {
        self.line_tokens.clear();
        let mut byte_offset = 0usize;
        for line_idx in 0..self.rope.len_lines() {
            let line = self.rope.line(line_idx).to_string();
            let tokens = tokenize_line(&line, byte_offset);
            self.line_tokens.push(tokens);
            byte_offset += line.len();
        }
        self.recompute_folds();
    }

    /// Retokenize a range of lines (inclusive start..=end)
    pub fn retokenize_range(&mut self, start_line: usize, end_line: usize) {
        if start_line >= self.rope.len_lines() { return; }
        let end_line = std::cmp::min(end_line, self.rope.len_lines() - 1);
        // Compute byte_offset for start_line
        let mut byte_offset = 0usize;
        for i in 0..start_line { byte_offset += self.rope.line(i).len_bytes(); }
        for li in start_line..=end_line {
            let line = self.rope.line(li).to_string();
            let tokens = tokenize_line(&line, byte_offset);
            if li < self.line_tokens.len() {
                self.line_tokens[li] = tokens;
            } else {
                self.line_tokens.push(tokens);
            }
            byte_offset += line.len();
        }
        self.recompute_folds();
    }

    /// Return a flat token list across all lines (useful for backends)
    pub fn tokenize_all(&self) -> Vec<TokenSpan> {
        let mut out = Vec::new();
        for line in &self.line_tokens {
            for t in line { out.push(t.clone()); }
        }
        out
    }

    fn recompute_folds(&mut self) {
        self.folds.clear();
        // Simple folding: match braces { } and also indent-based folds
        let mut stack: Vec<usize> = Vec::new();
        for (li, _line) in (0..self.rope.len_lines()).enumerate() {
            let s = self.rope.line(li).to_string();
            for ch in s.chars() {
                if ch == '{' {
                    stack.push(li);
                } else if ch == '}' {
                    if let Some(start) = stack.pop() {
                        if li > start + 0 {
                            self.folds.push((start, li));
                        }
                    }
                }
            }
        }
        // Note: indent-based folding could be added here as well
    }

    pub fn folds(&self) -> &Vec<(usize, usize)> { &self.folds }
}
