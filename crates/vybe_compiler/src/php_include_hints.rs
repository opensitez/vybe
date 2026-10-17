//! Cheap, speculative include discovery before the parent AST exists.
//! These hints never decide whether an include executes or what it loads.

use std::path::Path;

pub(crate) fn discover(source: &str, caller: &Path) -> Box<dyn crate::preparation::DeferredPreparationHints> {
    Box::new(IncludeHints::new(source, caller))
}

/// Optimized worker lookahead favors includes outside deferred function bodies.
/// A missed hint only loses speculation: execution still resolves and compiles
/// the requested file through the ordinary authoritative include path.
pub(crate) fn discover_immediate(source: &str, caller: &Path) -> Box<dyn crate::preparation::DeferredPreparationHints> {
    Box::new(IncludeHints::immediate(source, caller))
}

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

fn atom(source: &str, start: usize, caller: &Path,
    constant: &mut impl FnMut(&str) -> Option<String>,
) -> Option<(String, usize)> {
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
    let mut end = i;
    while bytes.get(end).is_some_and(|byte| ident(*byte)) { end += 1; }
    if end != i {
        return constant(&source[i..end]).map(|value| (value, end));
    }
    None
}


/// Only reads already-known string values. This never evaluates PHP code or
/// substitutes bindings into the prepared AST; the executing include remains
/// authoritative even if a value changes after this speculative lookup.
#[cfg(test)]
pub(crate) fn paths_with_constants(source: &str, caller: &Path,
    mut constant: impl FnMut(&str) -> Option<String>,
) -> Vec<String> {
    let bytes = source.as_bytes();
    let mut paths = Vec::new();
    let mut i = 0;
    // Tagged files can contain arbitrary HTML. Untagged units are accepted by
    // the frontend too, so scan them as code.
    let mut code = !source.contains("<?");
    while i < bytes.len() && paths.len() < 128 {
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
        let Some((mut path, end)) = atom(source, cursor, caller, &mut constant) else { continue; };
        cursor = trivia(bytes, end);
        let mut valid = true;
        while bytes.get(cursor) == Some(&b'.') {
            let Some((part, end)) = atom(source, cursor + 1, caller, &mut constant) else { valid = false; break; };
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

/// Request-owned symbolic hints. Resolved hints are consumed once; unresolved
/// constant operands can be retried at a later executing include boundary.
/// These expressions only schedule parsing, never execute or bind source code.
#[derive(Default)]
pub(crate) struct IncludeHints { pending: Vec<Vec<HintPart>> }
enum HintPart { Text(String), Constant(String) }
fn hint_atom(source: &str, start: usize, caller: &Path) -> Option<(HintPart, usize)> {
    if let Some((text, end)) = atom(source, start, caller, &mut |_| None) {
        return Some((HintPart::Text(text), end));
    }
    let bytes = source.as_bytes();
    let start = trivia(bytes, start);
    let mut end = start;
    while bytes.get(end).is_some_and(|byte| ident(*byte)) { end += 1; }
    (end != start).then(|| (HintPart::Constant(source[start..end].to_owned()), end))
}
impl IncludeHints {
    /// Consume parser-committed immediate operand text, rather than scouting
    /// the whole source again. Unsupported expressions remain runtime loads.
    pub(crate) fn from_operands(operands: impl IntoIterator<Item = String>, caller: &Path) -> Self {
        let mut pending = Vec::new();
        for source in operands {
            if pending.len() >= 128 { break; }
            let bytes = source.as_bytes();
            let mut cursor = trivia(bytes, 0);
            let parenthesized = bytes.get(cursor) == Some(&b'(');
            if parenthesized { cursor = trivia(bytes, cursor + 1); }
            let Some((first, end)) = hint_atom(&source, cursor, caller) else { continue; };
            let mut parts = vec![first];
            cursor = trivia(bytes, end);
            let mut valid = true;
            while bytes.get(cursor) == Some(&b'.') {
                let Some((part, end)) = hint_atom(&source, cursor + 1, caller) else { valid = false; break; };
                parts.push(part);
                cursor = trivia(bytes, end);
            }
            if parenthesized {
                valid &= bytes.get(cursor) == Some(&b')');
                cursor = trivia(bytes, cursor.saturating_add(1));
            }
            if valid && cursor == bytes.len() { pending.push(parts); }
        }
        Self { pending }
    }

    pub(crate) fn new(source: &str, caller: &Path) -> Self {
        Self::with_policy(source, caller, false)
    }

    // This is scheduling metadata only, not a reachability or execution claim.
    // Immediately invoked closures also fall back to actual include handling.
    pub(crate) fn immediate(source: &str, caller: &Path) -> Self {
        Self::with_policy(source, caller, true)
    }

    fn with_policy(source: &str, caller: &Path, immediate: bool) -> Self {
        let mut function_header = false;
        let mut function_depth = 0usize;
        let bytes = source.as_bytes();
        let mut hints = Vec::new();
        let mut i = 0;
        // Tagged files can contain arbitrary HTML. Untagged units are accepted by
        // the frontend too, so scan them as code.
        let mut code = !source.contains("<?");
        while i < bytes.len() {
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
            if immediate {
                if function_depth != 0 {
                    match bytes[i] { b'{' => function_depth += 1, b'}' => function_depth -= 1, _ => {} }
                    i += 1;
                    continue;
                }
                if function_header {
                    if bytes[i] == b'{' { function_header = false; function_depth = 1; i += 1; continue; }
                    if bytes[i] == b';' { function_header = false; }
                }
            }
            if !ident(bytes[i]) { i += 1; continue; }
            let start = i;
            while i < bytes.len() && ident(bytes[i]) { i += 1; }
            let token = &source[start..i];
            if immediate && token.eq_ignore_ascii_case("function") { function_header = true; continue; }
            if !["require", "require_once", "include", "include_once"].iter()
                .any(|keyword| token.eq_ignore_ascii_case(keyword)) { continue; }
            let mut cursor = trivia(bytes, i);
            let parenthesized = bytes.get(cursor) == Some(&b'(');
            if parenthesized { cursor += 1; }
            let Some((first, end)) = hint_atom(source, cursor, caller) else { continue; };
            cursor = trivia(bytes, end);
            let mut parts = vec![first];
            let mut valid = true;
            while bytes.get(cursor) == Some(&b'.') {
                let Some((part, end)) = hint_atom(source, cursor + 1, caller) else { valid = false; break; };
                parts.push(part);
                cursor = trivia(bytes, end);
            }
            if parenthesized {
                valid &= bytes.get(cursor) == Some(&b')');
                cursor = trivia(bytes, cursor.saturating_add(1));
            }
            // Reject continuations such as a ternary or a call. Missing delimiters
            // simply yield no hint; normal compilation still diagnoses the source.
            valid &= matches!(bytes.get(cursor), None | Some(b';' | b',' | b')'));
            if valid { hints.push(parts); }
        }
        Self { pending: hints }
    }

    pub(crate) fn take_ready_paths(&mut self, mut constant: impl FnMut(&str) -> Option<String>) -> Vec<String> {
        let mut paths = Vec::new();
        self.pending.retain_mut(|parts| {
            if paths.len() >= 128 { return true; }
            // Most literal includes have one operand. Transfer its owned path
            // rather than allocating a temporary values vector and copying it
            // into a second String. Unknown constants must remain retryable.
            if let [part] = parts.as_mut_slice() {
                let path = match part {
                    HintPart::Text(text) => std::mem::take(text),
                    HintPart::Constant(name) => match constant(name) {
                        Some(path) => path,
                        None => return true,
                    },
                };
                paths.push(path);
                return false;
            }
            let mut values = Vec::with_capacity(parts.len());
            for part in parts {
                match part {
                    HintPart::Text(text) => values.push(std::borrow::Cow::Borrowed(text.as_str())),
                    HintPart::Constant(name) => match constant(name) {
                        Some(value) => values.push(std::borrow::Cow::Owned(value)),
                        None => return true,
                    },
                }
            }
            let mut path = String::with_capacity(values.iter().map(|value| value.len()).sum());
            for value in values { path.push_str(&value); }
            paths.push(path);
            false
        });
        paths
    }
}

impl crate::preparation::DeferredPreparationHints for IncludeHints {
    fn take_ready_paths(&mut self, lookup: &mut dyn FnMut(&str) -> Option<String>) -> Vec<String> {
        IncludeHints::take_ready_paths(self, lookup)
    }
    fn is_empty(&self) -> bool { self.pending.is_empty() }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn committed_operands_match_scanner_hints_without_scanning_other_source() {
        let caller = Path::new("/site/main.php");
        let operands = ["'literal.php'", "(__DIR__ . '/magic.php')", "ABSPATH . WPINC . '/plugin.php'",
            "UNKNOWN . '/later.php'", "\"$base/a.php\"", "lookup()", "'prefix' . $name", "'x' ? 'y' : 'z'",
            "/*comment*/ (__FILE__ . '.extra')", "'a\\b.php'", "'unterminated", "'extra.php' trailing"];
        let source = operands.iter().map(|operand| format!("include {operand};")).collect::<String>();
        for mask in 0..8 {
            let lookup = |name: &str| match name {
                "ABSPATH" if mask & 1 != 0 => Some("/site/".to_owned()),
                "WPINC" if mask & 2 != 0 => Some("inc".to_owned()),
                "UNKNOWN" if mask & 4 != 0 => Some("other".to_owned()),
                _ => None,
            };
            // Compare individually: an unterminated operand must not consume
            // the next already-committed operand, unlike whole-source scouting.
            let expected = operands.iter().flat_map(|operand|
                paths_with_constants(&format!("include {operand};"), caller, lookup)).collect::<Vec<_>>();
            assert_eq!(IncludeHints::from_operands(operands.map(str::to_owned), caller)
                .take_ready_paths(lookup), expected, "mask={mask}; {source}");
        }
    }

    #[test]
    fn committed_operands_skip_unsupported_inputs_and_preserve_late_constant_retry() {
        let caller = Path::new("/site/main.php");
        let operands = std::iter::repeat_n("lookup()".to_owned(), 140)
            .chain(["LATE".to_owned(), "'ready.php'".to_owned()]);
        let mut hints = IncludeHints::from_operands(operands, caller);
        assert_eq!(hints.take_ready_paths(|_| None), ["ready.php"]);
        assert_eq!(hints.take_ready_paths(|name| (name == "LATE").then(|| "late.php".to_owned())), ["late.php"]);
        let mut hints = IncludeHints::from_operands((0..130).map(|i| format!("'{i}.php'")), caller);
        let ready = hints.take_ready_paths(|_| None);
        assert_eq!(ready.len(), 128);
        assert_eq!(ready.last().unwrap(), "127.php");
    }
    #[test]
    fn single_operand_hints_transfer_paths_and_retry_constants_in_order() {
        let mut hints = IncludeHints::new(
            "<?php include ''; require 'literal.php'; include LATE; require ROOT . '/child.php';",
            Path::new("/site/main.php"));
        let literal_pointer = match &hints.pending[1][0] {
            HintPart::Text(path) => path.as_ptr(),
            _ => panic!("expected literal hint"),
        };
        let mut lookups = Vec::new();
        let ready = hints.take_ready_paths(|name| { lookups.push(name.to_owned()); None });
        assert_eq!(ready, ["", "literal.php"]);
        assert_eq!(ready[1].as_ptr(), literal_pointer);
        assert_eq!(lookups, ["LATE", "ROOT"]);
        let mut supplied = Some("late.php".to_owned());
        let constant_pointer = supplied.as_ref().unwrap().as_ptr();
        let ready = hints.take_ready_paths(|name| match name {
            "LATE" => supplied.take(),
            "ROOT" => Some("/site".to_owned()),
            _ => None,
        });
        assert_eq!(ready, ["late.php", "/site/child.php"]);
        assert_eq!(ready[0].as_ptr(), constant_pointer);
        assert!(hints.pending.is_empty());
    }

    #[test]
    fn single_operand_hints_keep_the_batch_limit_and_pending_order() {
        let mut hints = IncludeHints { pending: (0..130)
            .map(|i| vec![HintPart::Text(format!("{i}.php"))]).collect() };
        let first = hints.take_ready_paths(|_| panic!("literal hint must not look up constants"));
        assert_eq!(first.len(), 128);
        assert_eq!(first.first().unwrap(), "0.php");
        assert_eq!(first.last().unwrap(), "127.php");
        assert_eq!(hints.take_ready_paths(|_| None), ["128.php", "129.php"]);
        assert!(hints.pending.is_empty());
    }
#[test] fn scheduling_only_function_filter(){
for s in ["<?php function cold(){require 'cold.php';} require 'hot.php';", "<?php class X {function cold(){if(true){require 'cold.php';}}} if(true){require 'hot.php';}", "<?php $f=function(){ require 'cold.php';}; require 'hot.php';", "<?php abstract class X {abstract function f();} require 'hot.php';", r#"<?php function cold(){ $x="} include 'fake.php';"; /* } */ require 'cold.php';} require 'hot.php';"#, "<?php function cold(){ ?> } <?php require 'cold.php';} require 'hot.php';"]{assert_eq!(IncludeHints::immediate(s, Path::new("/site/main.php")).take_ready_paths(|_| None),["hot.php"],"{s}");}
assert_eq!(IncludeHints::immediate("<?php require 'a.php'; if(true){include 'b.php';}", Path::new("/site/main.php")).take_ready_paths(|_| None),IncludeHints::new("<?php require 'a.php'; if(true){include 'b.php';}", Path::new("/site/main.php")).take_ready_paths(|_| None));
}

    #[test]
    fn symbolic_hints_retry_late_constants_without_republishing_ready_paths() {
        let mut hints = IncludeHints::new(
            "<?php require 'early.php'; require ABSPATH . WPINC . '/plugin.php'; include UNKNOWN . '/later.php';",
            Path::new("/site/settings.php"));
        assert_eq!(hints.take_ready_paths(|_| None), ["early.php"]);
        assert!(hints.take_ready_paths(|name| (name == "ABSPATH").then(|| "/site/".into())).is_empty());
        assert_eq!(hints.take_ready_paths(|name| match name {
            "ABSPATH" => Some("/site/".into()), "WPINC" => Some("wp-includes".into()), _ => None,
        }), ["/site/wp-includes/plugin.php"]);
        assert_eq!(hints.take_ready_paths(|_| Some("changed".into())), ["changed/later.php"]);
        assert!(hints.take_ready_paths(|_| Some("again".into())).is_empty());
    }

    #[test]
    fn symbolic_hints_preserve_legacy_paths_and_reject_dynamic_operands() {
        let caller = Path::new("/site/main.php");
        for source in [
            "<?php require ABSPATH . WPINC . '/x.php'; include UNKNOWN . '/y.php';",
            "<?php require /*c*/ (__DIR__ . '/a.php'); return include 'b.php';",
            "include 'html.php'; <?php // require 'comment.php';\n$x = \"include 'string.php';\"; require 'real.php'; ?> include 'html2.php';",
            "<?php include \"$base/a.php\"; include 'prefix' . $name; include 'x' ? 'y' : 'z'; include lookup() . '/y.php'; require 'real.php';",
            "<?php $x = <<<'TEXT'\ninclude 'heredoc.php';\nTEXT;\nrequire 'real.php';",
        ] {
            for mask in 0..8 {
                let constants = |name: &str| match name {
                    "ABSPATH" if mask & 1 != 0 => Some("/site/".to_owned()),
                    "WPINC" if mask & 2 != 0 => Some("inc".to_owned()),
                    "UNKNOWN" if mask & 4 != 0 => Some("other".to_owned()), _ => None,
                };
                assert_eq!(IncludeHints::new(source, caller).take_ready_paths(constants),
                    paths_with_constants(source, caller, constants), "{source}; mask={mask}");
            }
        }
    }

    fn paths(source: &str, caller: &Path) -> Vec<String> {
        paths_with_constants(source, caller, |_| None)
    }

    #[test]
    fn uses_existing_constants_for_wordpress_style_paths_without_evaluation() {
        let source = "<?php require ABSPATH . WPINC . '/plugin.php'; require UNKNOWN . '/x.php'; require lookup() . '/y.php';";
        let paths = paths_with_constants(source, Path::new("/site/main.php"), |name| match name {
            "ABSPATH" => Some("/site/".into()),
            "WPINC" => Some("wp-includes".into()),
            _ => None,
        });
        assert_eq!(paths, ["/site/wp-includes/plugin.php"]);
    }

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
