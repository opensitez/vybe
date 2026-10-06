//! Cheap, speculative include discovery before the parent AST exists.
//! These hints never decide whether an include executes or what it loads.

use std::path::Path;

fn ident(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_' || byte >= 128
}

fn trivia(bytes: &[u8], mut i: usize) -> usize {
    loop {
        while bytes.get(i).is_some_and(u8::is_ascii_whitespace) { i += 1; }
        if bytes.get(i..i + 2) == Some(b"/*") {
            i += 2;
            while i < bytes.len() && bytes.get(i..i + 2) != Some(b"*/") { i += 1; }
            i = (i + 2).min(bytes.len());
        } else if bytes.get(i..i + 2) == Some(b"//") || bytes.get(i) == Some(&b'#') {
            while i < bytes.len() && bytes[i] != b'\n' { i += 1; }
        } else { return i; }
    }
}

fn quoted(bytes: &[u8], start: usize) -> usize {
    let quote = bytes[start];
    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] == b'\\' { i = (i + 2).min(bytes.len()); }
        else if bytes[i] == quote { return i + 1; }
        else { i += 1; }
    }
    i
}

fn atom(source: &str, start: usize, caller: &Path) -> Option<(String, usize)> {
    let bytes = source.as_bytes();
    let i = trivia(bytes, start);
    if matches!(bytes.get(i), Some(b'\'' | b'"')) {
        let end = quoted(bytes, i);
        if end <= i + 1 || bytes.get(end - 1) != bytes.get(i) { return None; }
        let value = &source[i + 1..end - 1];
        // Escapes and interpolation need PHP's parser. Leave them to the AST
        // pass rather than guessing their values in the lookahead scanner.
        if value.contains('\\') || (bytes[i] == b'"' && value.contains('$')) { return None; }
        return Some((value.to_owned(), end));
    }
    for (token, path) in [("__DIR__", caller.parent()?), ("__FILE__", caller)] {
        if source[i..].starts_with(token) && !bytes.get(i + token.len()).is_some_and(|b| ident(*b)) {
            return Some((path.to_string_lossy().into_owned(), i + token.len()));
        }
    }
    None
}

pub(crate) fn paths(source: &str, caller: &Path) -> Vec<String> {
    let bytes = source.as_bytes();
    let mut paths = Vec::new();
    let mut i = 0;
    // Tagged files can contain arbitrary HTML. Untagged units are accepted by
    // the frontend too, so scan them as code.
    let mut code = !source.contains("<?");
    while i < bytes.len() && paths.len() < 16 {
        if !code {
            if bytes.get(i..i + 2) == Some(b"<?") { code = true; i += 2; }
            else { i += 1; }
            continue;
        }
        if bytes.get(i..i + 2) == Some(b"?>") { code = false; i += 2; continue; }
        let next = trivia(bytes, i);
        if next != i { i = next; continue; }
        if let Some(end) = crate::bundle::skip_php_heredoc(bytes, i) { i = end; continue; }
        if matches!(bytes[i], b'\'' | b'"' | b'`') { i = quoted(bytes, i); continue; }
        if !ident(bytes[i]) { i += 1; continue; }
        let start = i;
        while i < bytes.len() && ident(bytes[i]) { i += 1; }
        let token = &source[start..i];
        if !["require", "require_once", "include", "include_once"].iter()
            .any(|keyword| token.eq_ignore_ascii_case(keyword)) { continue; }
        let mut cursor = trivia(bytes, i);
        let parenthesized = bytes.get(cursor) == Some(&b'(');
        if parenthesized { cursor += 1; }
        let Some((mut path, end)) = atom(source, cursor, caller) else { continue; };
        cursor = trivia(bytes, end);
        let mut valid = true;
        while bytes.get(cursor) == Some(&b'.') {
            let Some((part, end)) = atom(source, cursor + 1, caller) else { valid = false; break; };
            path.push_str(&part);
            cursor = trivia(bytes, end);
        }
        if parenthesized {
            valid &= bytes.get(cursor) == Some(&b')');
            cursor = trivia(bytes, cursor.saturating_add(1));
        }
        // Reject continuations such as a ternary or a call. Missing delimiters
        // simply yield no hint; normal compilation still diagnoses the source.
        valid &= matches!(bytes.get(cursor), None | Some(b';' | b',' | b')'));
        if valid { paths.push(path); }
    }
    paths
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn finds_literals_and_magic_paths_without_waiting_for_an_ast() {
        let source = "<?php require /*x*/ (__DIR__ . '/a.php'); return include 'b.php'; INCLUDE_ONCE \"c.php\";";
        assert_eq!(paths(source, Path::new("/site/main.php")), ["/site/a.php", "b.php", "c.php"]);
    }

    #[test]
    fn ignores_html_comments_strings_heredocs_and_dynamic_expressions() {
        let source = "include 'html.php'; <?php // include 'comment.php';\n\
            $x = \"include 'string.php';\"; $x = <<<'TEXT'\ninclude 'heredoc.php';\nTEXT;\n\
            include \"$base/a.php\"; include 'prefix' . $name; include 'x' ? 'y' : 'z';\n\
            require 'real.php'; ?> include 'html2.php';";
        assert_eq!(paths(source, Path::new("/site/main.php")), ["real.php"]);
    }
}
