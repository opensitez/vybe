//! Scanner for the Rust literals that hold test sources.
//!
//! Test bodies reach us as Rust string literals, and they come in two forms
//! with very different risk: `r#"..."#` is verbatim (Python/PHP use it), while
//! `"..."` carries backslash escapes (Go/C/Lua use it) and in Go's case nests
//! backtick struct tags inside those escapes — `` `json:\",omitempty\"` ``.
//! Getting the unescape wrong there corrupts sources silently, so `escape`
//! below is the inverse of this and every extraction is round-tripped through
//! both before it is written out.

/// Read one Rust string literal starting at `at` (which must index the opening
/// `"`, or the `r` of a raw literal). Returns the decoded text and the index
/// just past the closing delimiter.
pub fn scan(src: &[u8], at: usize) -> anyhow::Result<(String, usize)> {
    if src.get(at) == Some(&b'r') {
        return scan_raw(src, at);
    }
    if src.get(at) != Some(&b'"') {
        anyhow::bail!("expected a string literal at byte {at}");
    }

    let mut out = String::new();
    let mut i = at + 1;
    while i < src.len() {
        match src[i] {
            b'"' => return Ok((out, i + 1)),
            b'\\' => {
                let (ch, next) = scan_escape(src, i + 1)?;
                if let Some(ch) = ch {
                    out.push(ch);
                }
                i = next;
            }
            _ => {
                // Multi-byte UTF-8 passes through whole; the corpus is full of
                // "привет" and "😀" and splitting a code point here would be
                // the same class of bug as the surrogate one in json.rs.
                let len = utf8_len(src[i]);
                out.push_str(std::str::from_utf8(&src[i..i + len])?);
                i += len;
            }
        }
    }
    anyhow::bail!("unterminated string literal starting at byte {at}")
}

/// One escape sequence, positioned just after the backslash. `Ok((None, _))`
/// means the escape produced no character — the `\<newline>` line continuation,
/// which swallows the newline and the indentation that follows it.
fn scan_escape(src: &[u8], at: usize) -> anyhow::Result<(Option<char>, usize)> {
    let Some(&tag) = src.get(at) else {
        anyhow::bail!("trailing backslash at byte {at}");
    };
    Ok(match tag {
        b'n' => (Some('\n'), at + 1),
        b'r' => (Some('\r'), at + 1),
        b't' => (Some('\t'), at + 1),
        b'0' => (Some('\0'), at + 1),
        b'\\' => (Some('\\'), at + 1),
        b'"' => (Some('"'), at + 1),
        b'\'' => (Some('\''), at + 1),
        b'x' => {
            let hex = std::str::from_utf8(&src[at + 1..at + 3])?;
            let code = u8::from_str_radix(hex, 16)?;
            (Some(code as char), at + 3)
        }
        b'u' => {
            let open = at + 1;
            if src.get(open) != Some(&b'{') {
                anyhow::bail!("\\u escape without a brace at byte {at}");
            }
            let close = (open..src.len())
                .find(|&i| src[i] == b'}')
                .ok_or_else(|| anyhow::anyhow!("unterminated \\u{{...}} at byte {at}"))?;
            let hex = std::str::from_utf8(&src[open + 1..close])?;
            let code = u32::from_str_radix(hex, 16)?;
            let ch = char::from_u32(code)
                .ok_or_else(|| anyhow::anyhow!("\\u{{{hex}}} is not a scalar value"))?;
            (Some(ch), close + 1)
        }
        b'\n' => {
            let mut i = at + 1;
            while matches!(src.get(i), Some(b' ' | b'\t' | b'\r' | b'\n')) {
                i += 1;
            }
            (None, i)
        }
        other => anyhow::bail!("unknown escape \\{} at byte {at}", other as char) })
}

fn scan_raw(src: &[u8], at: usize) -> anyhow::Result<(String, usize)> {
    let mut hashes = 0usize;
    let mut i = at + 1;
    while src.get(i) == Some(&b'#') {
        hashes += 1;
        i += 1;
    }
    if src.get(i) != Some(&b'"') {
        anyhow::bail!("malformed raw string literal at byte {at}");
    }
    i += 1;
    let start = i;

    let mut terminator = String::from("\"");
    terminator.push_str(&"#".repeat(hashes));
    let term = terminator.as_bytes();

    while i + term.len() <= src.len() {
        if &src[i..i + term.len()] == term {
            return Ok((String::from_utf8(src[start..i].to_vec())?, i + term.len()));
        }
        i += 1;
    }
    anyhow::bail!("unterminated raw string literal starting at byte {at}")
}

/// The inverse of `scan` for a non-raw literal. Only used to prove a decode was
/// lossless, never to write Rust.
pub fn escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            '\0' => out.push_str("\\0"),
            _ => out.push(ch) }
    }
    out
}

fn utf8_len(byte: u8) -> usize {
    match byte {
        0x00..=0x7F => 1,
        0xC0..=0xDF => 2,
        0xE0..=0xEF => 3,
        _ => 4 }
}

/// Skip whitespace and `//` / `/* */` comments starting at `at`.
pub fn skip_trivia(src: &[u8], mut at: usize) -> usize {
    loop {
        while at < src.len() && src[at].is_ascii_whitespace() {
            at += 1;
        }
        if src[at..].starts_with(b"//") {
            while at < src.len() && src[at] != b'\n' {
                at += 1;
            }
            continue;
        }
        if src[at..].starts_with(b"/*") {
            at += 2;
            while at + 1 < src.len() && &src[at..at + 2] != b"*/" {
                at += 1;
            }
            at = (at + 2).min(src.len());
            continue;
        }
        return at;
    }
}
