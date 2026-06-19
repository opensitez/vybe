//! PHP walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//! Once this returns a `Module`, the rest of the compilation pipeline
//! (compile_class / compile_expression / etc.) is shared with every
//! other vybex language and works without any PHP-specific knowledge.
//!
//! ## Notes on PHP semantics that the walker normalises
//!
//! - **`$variable`** vs **bare identifier**. PHP distinguishes them at
//!   the lexical level: `$x` is a variable, `x` is a function name or a
//!   constant. The walker emits `Ident` for both kinds — for variables
//!   we strip the leading `$` so the canonical AST identifier matches
//!   what every other language uses.
//!
//! - **`echo` and `print`** become `StmtKind::Echo(...)` directly, which
//!   the compiler routes through `compiler_common::io::emit_print`.
//!
//! - **`$obj->method()`** is `Call { callee: Member { object, field } }`
//!   exactly like JS `obj.method()`. PHP `?->` becomes `Member {
//!   null_safe: true }`.
//!
//! - **`Class::method()` / `Class::CONST`** uses `ExprKind::StaticAccess`
//!   which the compiler treats as a struct_get on the class global.
//!
//! - **`use Foo\Bar;`** and **`namespace Foo\Bar;`** are parsed but
//!   discarded — the compiler treats every name as a flat global. PHP
//!   namespaces are mostly cosmetic for our purposes.
//!
//! - **Type hints** are parsed and discarded. We don't type-check.
//!
//! - **Promoted constructor parameters** (PHP 8): `public int $foo` in
//!   the constructor parameter list. The walker emits the param AND
//!   synthesises a property + an assignment in the body so the
//!   downstream compiler doesn't need to know about the promotion.
//!
//! - **`<?php` open tag**: stripped at the grammar level (`open_tag` is
//!   silent). User scripts may or may not have it.

use super::{PhpParser, Rule};
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;
use std::cell::RefCell;

// Class context for `self::` resolution. PHP `self::X` inside a method
// refers to the enclosing class (NOT the runtime instance) — it's a
// compile-time-known reference. The walker pushes the current class
// name before walking class members and pops it after, so when
// `Rule::kw_self` is reached we can rewrite `self` to the class name
// directly and avoid the runtime "STRUCT_GET on $this" path that
// can't reach class-level constants/static members.
thread_local! {
    static CLASS_STACK: RefCell<Vec<String>> = const { RefCell::new(Vec::new()) };
    // Tracks `use TraitName;` per class. Walker captures the trait name
    // when it sees `use_trait` inside a class member; the post-pass in
    // `parse()` reads this to copy trait members into the using class.
    // Reset at the start of each `parse()` call.
    static TRAIT_USAGES: RefCell<std::collections::HashMap<String, Vec<String>>> =
        RefCell::new(std::collections::HashMap::new());
    // Tracks `use Trait { method as alias; }` adaptations. Map key is
    // the using-class name, value is a list of (source_method, alias)
    // pairs. Reset alongside TRAIT_USAGES.
    // Each entry: (source_trait_name | "" if unqualified, method_name, alias_name).
    static TRAIT_ALIASES: RefCell<std::collections::HashMap<String, Vec<(String, String, String)>>> =
        RefCell::new(std::collections::HashMap::new());
    // Monotonic counter for unique synthetic temp variable names. The
    // postfix `$x++` walker rewrite needs `(tmp = $x, $x = inc($x), tmp)`
    // where `tmp` must be unique per use site to avoid collisions in
    // expressions like `$a++ + $b++` (without uniqueness, both writes
    // would clobber the same global temp).
    static TMP_COUNTER: RefCell<u32> = const { RefCell::new(0) };
    // Suppression flag for the magic-get/`__get` rewrite: when
    // `walk_assignment` is processing its LHS, the outermost
    // `property_access_op` is the assignment TARGET, not a read, and
    // should not be wrapped in the magic-get ternary (the wrapped
    // expression isn't an l-value). Walker increments before walking
    // the LHS and decrements after; `apply_postfix` peeks at the
    // depth to decide whether to skip the wrap on the LAST chain op.
    static ASSIGN_LHS_DEPTH: RefCell<u32> = const { RefCell::new(0) };
    static LINE_STARTS: RefCell<Vec<usize>> = const { RefCell::new(Vec::new()) };
}

const PHP_LITERAL_OPEN_MASK: &str = "\u{E000}\u{E001}";
const PHP_LITERAL_CLOSE_MASK: &str = "\u{E001}\u{E000}";

fn unmask_php_literal_tags(text: &str) -> String {
    text.replace(PHP_LITERAL_OPEN_MASK, "<?")
        .replace(PHP_LITERAL_CLOSE_MASK, "?>")
}

#[allow(dead_code)]
fn parse_php_heredoc_header(bytes: &[u8], start: usize) -> Option<(usize, Vec<u8>)> {
    if start + 3 >= bytes.len()
        || bytes[start] != b'<'
        || bytes[start + 1] != b'<'
        || bytes[start + 2] != b'<'
    {
        return None;
    }

    let mut index = start + 3;
    while index < bytes.len() && matches!(bytes[index], b' ' | b'\t') {
        index += 1;
    }

    let quote = match bytes.get(index).copied() {
        Some(b'\'') | Some(b'"') => {
            let q = bytes[index];
            index += 1;
            Some(q)
        }
        _ => None,
    };

    let tag_start = index;
    if !matches!(
        bytes.get(index).copied(),
        Some(b'A'..=b'Z' | b'a'..=b'z' | b'_')
    ) {
        return None;
    }
    index += 1;
    while index < bytes.len()
        && matches!(bytes[index], b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_')
    {
        index += 1;
    }
    let tag = bytes[tag_start..index].to_vec();

    if let Some(q) = quote {
        if bytes.get(index).copied() != Some(q) {
            return None;
        }
        index += 1;
    }

    while index < bytes.len() {
        match bytes[index] {
            b'\n' => {
                index += 1;
                break;
            }
            b'\r' => {
                index += 1;
                if bytes.get(index).copied() == Some(b'\n') {
                    index += 1;
                }
                break;
            }
            _ => index += 1,
        }
    }

    Some((index, tag))
}

#[allow(dead_code)]
fn mask_php_literal_tag_sequences(source: &str) -> String {
    enum ScanState {
        Normal,
        SingleQuote,
        DoubleQuote,
        LineComment,
        BlockComment,
        Heredoc { tag: Vec<u8>, at_line_start: bool },
    }

    let bytes = source.as_bytes();
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len());
    let mut index = 0usize;
    let mut state = ScanState::Normal;

    while index < bytes.len() {
        match &mut state {
            ScanState::Normal => {
                if let Some((after_header, tag)) = parse_php_heredoc_header(bytes, index) {
                    out.extend_from_slice(&bytes[index..after_header]);
                    index = after_header;
                    state = ScanState::Heredoc {
                        tag,
                        at_line_start: true,
                    };
                    continue;
                }
                if bytes[index] == b'/' && index + 1 < bytes.len() && bytes[index + 1] == b'/' {
                    out.extend_from_slice(b"//");
                    index += 2;
                    state = ScanState::LineComment;
                    continue;
                }
                if bytes[index] == b'#' {
                    out.push(bytes[index]);
                    index += 1;
                    state = ScanState::LineComment;
                    continue;
                }
                if bytes[index] == b'/' && index + 1 < bytes.len() && bytes[index + 1] == b'*' {
                    out.extend_from_slice(b"/*");
                    index += 2;
                    state = ScanState::BlockComment;
                    continue;
                }
                if bytes[index] == b'\'' {
                    out.push(bytes[index]);
                    index += 1;
                    state = ScanState::SingleQuote;
                    continue;
                }
                if bytes[index] == b'"' {
                    out.push(bytes[index]);
                    index += 1;
                    state = ScanState::DoubleQuote;
                    continue;
                }
                out.push(bytes[index]);
                index += 1;
            }
            ScanState::SingleQuote => {
                if index + 1 < bytes.len() && bytes[index] == b'<' && bytes[index + 1] == b'?' {
                    out.extend_from_slice(PHP_LITERAL_OPEN_MASK.as_bytes());
                    index += 2;
                    continue;
                }
                if index + 1 < bytes.len() && bytes[index] == b'?' && bytes[index + 1] == b'>' {
                    out.extend_from_slice(PHP_LITERAL_CLOSE_MASK.as_bytes());
                    index += 2;
                    continue;
                }
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    out.extend_from_slice(&bytes[index..index + 2]);
                    index += 2;
                    continue;
                }
                out.push(bytes[index]);
                if bytes[index] == b'\'' {
                    state = ScanState::Normal;
                }
                index += 1;
            }
            ScanState::DoubleQuote => {
                if index + 1 < bytes.len() && bytes[index] == b'<' && bytes[index + 1] == b'?' {
                    out.extend_from_slice(PHP_LITERAL_OPEN_MASK.as_bytes());
                    index += 2;
                    continue;
                }
                if index + 1 < bytes.len() && bytes[index] == b'?' && bytes[index + 1] == b'>' {
                    out.extend_from_slice(PHP_LITERAL_CLOSE_MASK.as_bytes());
                    index += 2;
                    continue;
                }
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    out.extend_from_slice(&bytes[index..index + 2]);
                    index += 2;
                    continue;
                }
                out.push(bytes[index]);
                if bytes[index] == b'"' {
                    state = ScanState::Normal;
                }
                index += 1;
            }
            ScanState::LineComment => {
                out.push(bytes[index]);
                let is_newline = bytes[index] == b'\n';
                index += 1;
                if is_newline {
                    state = ScanState::Normal;
                }
            }
            ScanState::BlockComment => {
                out.push(bytes[index]);
                if bytes[index] == b'*' && index + 1 < bytes.len() && bytes[index + 1] == b'/' {
                    out.push(bytes[index + 1]);
                    index += 2;
                    state = ScanState::Normal;
                    continue;
                }
                index += 1;
            }
            ScanState::Heredoc { tag, at_line_start } => {
                if *at_line_start {
                    let mut check = index;
                    while check < bytes.len() && matches!(bytes[check], b' ' | b'\t') {
                        check += 1;
                    }
                    if bytes[check..].starts_with(tag) {
                        let mut after = check + tag.len();
                        if after == bytes.len() || matches!(bytes[after], b';' | b'\n' | b'\r') {
                            if bytes.get(after).copied() == Some(b';') {
                                after += 1;
                            }
                            while after < bytes.len() {
                                match bytes[after] {
                                    b'\n' => {
                                        after += 1;
                                        break;
                                    }
                                    b'\r' => {
                                        after += 1;
                                        if bytes.get(after).copied() == Some(b'\n') {
                                            after += 1;
                                        }
                                        break;
                                    }
                                    _ => after += 1,
                                }
                            }
                            out.extend_from_slice(&bytes[index..after]);
                            index = after;
                            state = ScanState::Normal;
                            continue;
                        }
                    }
                }
                if index + 1 < bytes.len() && bytes[index] == b'<' && bytes[index + 1] == b'?' {
                    out.extend_from_slice(PHP_LITERAL_OPEN_MASK.as_bytes());
                    index += 2;
                    *at_line_start = false;
                    continue;
                }
                if index + 1 < bytes.len() && bytes[index] == b'?' && bytes[index + 1] == b'>' {
                    out.extend_from_slice(PHP_LITERAL_CLOSE_MASK.as_bytes());
                    index += 2;
                    *at_line_start = false;
                    continue;
                }
                let byte = bytes[index];
                out.push(byte);
                index += 1;
                *at_line_start = byte == b'\n';
            }
        }
    }

    String::from_utf8(out).unwrap_or_else(|_| source.to_string())
}

fn build_line_starts(source: &str) -> Vec<usize> {
    let mut starts = Vec::with_capacity(source.len() / 32 + 2);
    starts.push(0);
    for (idx, byte) in source.bytes().enumerate() {
        if byte == b'\n' {
            starts.push(idx + 1);
        }
    }
    starts
}

fn offset_to_line_col(offset: usize, line_starts: &[usize]) -> (u32, u32) {
    let line_index = match line_starts.binary_search(&offset) {
        Ok(index) => index,
        Err(next_index) => next_index.saturating_sub(1),
    };
    let line_start = line_starts.get(line_index).copied().unwrap_or(0);
    let line = (line_index + 1) as u32;
    let col = (offset.saturating_sub(line_start) + 1) as u32;
    (line, col)
}

fn next_tmp_name(prefix: &str) -> String {
    TMP_COUNTER.with(|c| {
        let mut n = c.borrow_mut();
        *n += 1;
        format!("__php_{}_{}", prefix, *n)
    })
}

fn walk_expression_as_assign_target(pair: Pair<Rule>) -> Result<Expression, String> {
    ASSIGN_LHS_DEPTH.with(|d| *d.borrow_mut() += 1);
    let walked = walk_expression(pair);
    ASSIGN_LHS_DEPTH.with(|d| {
        let mut bd = d.borrow_mut();
        *bd = bd.saturating_sub(1);
    });
    walked
}

fn postfix_rule_kind(pair: &Pair<Rule>) -> Option<Rule> {
    if matches!(pair.as_rule(), Rule::postfix_op) {
        pair.clone()
            .into_inner()
            .next()
            .map(|inner| inner.as_rule())
    } else {
        Some(pair.as_rule())
    }
}

fn push_class_context(name: &str) {
    CLASS_STACK.with(|s| s.borrow_mut().push(name.to_string()));
}

fn pop_class_context() {
    CLASS_STACK.with(|s| {
        s.borrow_mut().pop();
    });
}

fn current_class_name() -> Option<String> {
    CLASS_STACK.with(|s| s.borrow().last().cloned())
}

#[allow(dead_code)]
enum MixedPhpSegment<'a> {
    Html(&'a str),
    Code { code: &'a str, has_close_tag: bool },
    Echo { expr: &'a str, has_close_tag: bool },
}

fn append_html_echo(out: &mut String, html: &str) {
    if html.is_empty() {
        return;
    }
    out.push_str("echo '");
    for ch in html.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '\'' => out.push_str("\\'"),
            _ => out.push(ch),
        }
    }
    out.push_str("';\n");
}

fn code_block_needs_terminator(code: &str) -> bool {
    let trimmed = code.trim_end();
    let Some(last) = trimmed.chars().last() else {
        return false;
    };
    !matches!(last, ';' | '{' | '}' | ':')
}

fn normalize_mixed_php_source(source: &str) -> String {
    let mut out = String::new();
    for segment in split_mixed_php_source(source) {
        match segment {
            MixedPhpSegment::Html(text) => append_html_echo(&mut out, text),
            MixedPhpSegment::Echo { expr, .. } => {
                let expr = expr.trim();
                if !expr.is_empty() {
                    out.push_str("echo ");
                    out.push_str(expr);
                    out.push_str(";\n");
                }
            }
            MixedPhpSegment::Code {
                code,
                has_close_tag,
            } => {
                out.push_str(code);
                if has_close_tag && code_block_needs_terminator(code) {
                    out.push(';');
                }
                if !out.ends_with('\n') {
                    out.push('\n');
                }
            }
        }
    }
    out
}

pub(crate) fn normalize_source_for_parser(source: &str) -> String {
    normalize_mixed_php_source(source)
}

fn split_mixed_php_source(source: &str) -> Vec<MixedPhpSegment<'_>> {
    let mut segments = Vec::new();
    let mut cursor = 0usize;

    while let Some(open_rel) = source[cursor..].find("<?") {
        let open = cursor + open_rel;
        if open > cursor {
            segments.push(MixedPhpSegment::Html(&source[cursor..open]));
        }

        let is_echo = source[open..].starts_with("<?=");
        let code_start = if is_echo {
            open + 3
        } else if source[open..].starts_with("<?php") {
            open + 5
        } else {
            open + 2
        };
        let close = find_php_close_tag(source, code_start).unwrap_or(source.len());
        let has_close_tag = close < source.len();
        let code = &source[code_start..close];
        if is_echo {
            segments.push(MixedPhpSegment::Echo {
                expr: code,
                has_close_tag,
            });
        } else {
            segments.push(MixedPhpSegment::Code {
                code,
                has_close_tag,
            });
        }
        cursor = if has_close_tag {
            (close + 2).min(source.len())
        } else {
            close
        };
    }

    if cursor < source.len() {
        segments.push(MixedPhpSegment::Html(&source[cursor..]));
    }

    segments
}

fn skip_php_heredoc(bytes: &[u8], start: usize) -> Option<usize> {
    if start + 3 >= bytes.len()
        || bytes[start] != b'<'
        || bytes[start + 1] != b'<'
        || bytes[start + 2] != b'<'
    {
        return None;
    }

    let mut index = start + 3;
    while index < bytes.len() && matches!(bytes[index], b' ' | b'\t') {
        index += 1;
    }

    let quote = match bytes.get(index).copied() {
        Some(b'\'') | Some(b'"') => {
            let q = bytes[index];
            index += 1;
            Some(q)
        }
        _ => None,
    };

    let tag_start = index;
    if !matches!(
        bytes.get(index).copied(),
        Some(b'A'..=b'Z' | b'a'..=b'z' | b'_')
    ) {
        return None;
    }
    index += 1;
    while index < bytes.len()
        && matches!(bytes[index], b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_')
    {
        index += 1;
    }
    let tag = &bytes[tag_start..index];

    if let Some(q) = quote {
        if bytes.get(index).copied() != Some(q) {
            return None;
        }
        index += 1;
    }

    while index < bytes.len() {
        match bytes[index] {
            b'\n' => {
                index += 1;
                break;
            }
            b'\r' => {
                index += 1;
                if bytes.get(index).copied() == Some(b'\n') {
                    index += 1;
                }
                break;
            }
            _ => index += 1,
        }
    }

    while index < bytes.len() {
        let line_start = index;
        while index < bytes.len() && matches!(bytes[index], b' ' | b'\t') {
            index += 1;
        }

        if bytes[index..].starts_with(tag) {
            let mut after = index + tag.len();
            if after == bytes.len() || matches!(bytes[after], b';' | b'\n' | b'\r') {
                if bytes.get(after).copied() == Some(b';') {
                    after += 1;
                }
                while after < bytes.len() {
                    match bytes[after] {
                        b'\n' => {
                            after += 1;
                            break;
                        }
                        b'\r' => {
                            after += 1;
                            if bytes.get(after).copied() == Some(b'\n') {
                                after += 1;
                            }
                            break;
                        }
                        _ => after += 1,
                    }
                }
                return Some(after);
            }
        }

        index = line_start;
        while index < bytes.len() {
            match bytes[index] {
                b'\n' => {
                    index += 1;
                    break;
                }
                b'\r' => {
                    index += 1;
                    if bytes.get(index).copied() == Some(b'\n') {
                        index += 1;
                    }
                    break;
                }
                _ => index += 1,
            }
        }
    }

    Some(bytes.len())
}

fn find_php_close_tag(source: &str, start: usize) -> Option<usize> {
    #[derive(Copy, Clone, Eq, PartialEq)]
    enum ScanState {
        Normal,
        SingleQuote,
        DoubleQuote,
        LineComment,
        BlockComment,
    }

    let bytes = source.as_bytes();
    let mut index = start;
    let mut state = ScanState::Normal;

    while index + 1 < bytes.len() {
        match state {
            ScanState::Normal => {
                if let Some(next_index) = skip_php_heredoc(bytes, index) {
                    index = next_index;
                    continue;
                }
                if bytes[index] == b'?' && bytes[index + 1] == b'>' {
                    return Some(index);
                }
                if bytes[index] == b'\'' {
                    state = ScanState::SingleQuote;
                } else if bytes[index] == b'"' {
                    state = ScanState::DoubleQuote;
                } else if bytes[index] == b'#' {
                    state = ScanState::LineComment;
                } else if bytes[index] == b'/' && bytes[index + 1] == b'/' {
                    state = ScanState::LineComment;
                    index += 1;
                } else if bytes[index] == b'/' && bytes[index + 1] == b'*' {
                    state = ScanState::BlockComment;
                    index += 1;
                }
            }
            ScanState::SingleQuote => {
                if bytes[index] == b'\\' {
                    index += 1;
                } else if bytes[index] == b'\'' {
                    state = ScanState::Normal;
                }
            }
            ScanState::DoubleQuote => {
                if bytes[index] == b'\\' {
                    index += 1;
                } else if bytes[index] == b'"' {
                    state = ScanState::Normal;
                }
            }
            ScanState::LineComment => {
                if bytes[index] == b'?' && bytes[index + 1] == b'>' {
                    return Some(index);
                }
                if bytes[index] == b'\n' {
                    state = ScanState::Normal;
                }
            }
            ScanState::BlockComment => {
                if bytes[index] == b'*' && bytes[index + 1] == b'/' {
                    state = ScanState::Normal;
                    index += 1;
                }
            }
        }
        index += 1;
    }

    None
}

/// Returns true for `kw_*` token rules. Pest preserves atomic rule
/// nodes as siblings inside their parent rule's parse tree, so without
/// this filter the keyword tokens leak into walker positional indexing
/// (e.g. `if (...)` would land `kw_if` as the first child of
/// `if_statement` and walk_if would try to parse it as an expression).
fn is_kw(r: Rule) -> bool {
    matches!(
        r,
        Rule::kw_if
            | Rule::kw_elseif
            | Rule::kw_else_if
            | Rule::kw_else
            | Rule::kw_while
            | Rule::kw_do
            | Rule::kw_for
            | Rule::kw_foreach
            | Rule::kw_as
            | Rule::kw_switch
            | Rule::kw_case
            | Rule::kw_default
            | Rule::kw_break
            | Rule::kw_continue
            | Rule::kw_return
            | Rule::kw_function
            | Rule::kw_class
            | Rule::kw_extends
            | Rule::kw_implements
            | Rule::kw_interface
            | Rule::kw_trait
            | Rule::kw_enum
            | Rule::kw_new
            | Rule::kw_clone
            | Rule::kw_echo
            | Rule::kw_print
            | Rule::kw_include
            | Rule::kw_include_once
            | Rule::kw_require
            | Rule::kw_require_once
            | Rule::kw_null
            | Rule::kw_true
            | Rule::kw_false
            | Rule::kw_instanceof
            | Rule::kw_throw
            | Rule::kw_try
            | Rule::kw_catch
            | Rule::kw_finally
            | Rule::kw_static
            | Rule::kw_public
            | Rule::kw_private
            | Rule::kw_protected
            | Rule::kw_abstract
            | Rule::kw_final
            | Rule::kw_const
            | Rule::kw_match
            | Rule::kw_fn
            | Rule::kw_use
            | Rule::kw_namespace
            | Rule::kw_yield_from
            | Rule::kw_yield
            | Rule::kw_list
            | Rule::kw_global
            | Rule::kw_readonly
            | Rule::kw_and
            | Rule::kw_or
            | Rule::kw_xor
            | Rule::kw_self
            | Rule::kw_parent
            | Rule::kw_isset
            | Rule::kw_empty
            | Rule::kw_unset
            | Rule::kw_endif
            | Rule::kw_endwhile
            | Rule::kw_endfor
            | Rule::kw_endforeach
            | Rule::kw_endswitch
            | Rule::kw_insteadof
    )
}

/// `pair.into_inner()` with `kw_*` siblings stripped — use this in any
/// walker that does positional indexing on a rule body that includes
/// keywords.
fn inner_nokw(pair: Pair<Rule>) -> std::vec::IntoIter<Pair<Rule>> {
    let kept: Vec<Pair<Rule>> = pair.into_inner().filter(|p| !is_kw(p.as_rule())).collect();
    kept.into_iter()
}

fn collect_program_body(
    program: Pair<Rule>,
    body: &mut Vec<Statement>,
    interface_names: &mut std::collections::HashSet<String>,
    trait_names: &mut std::collections::HashSet<String>,
) -> Result<(), String> {
    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::EOI => continue,
            Rule::inline_html => {
                let text = pair.as_str();
                if !text.is_empty() {
                    body.push(Statement::new(StmtKind::Echo(vec![Expression::new(
                        ExprKind::Lit(Literal::Str(text.to_string())),
                    )])));
                }
            }
            Rule::php_code_segment_closed | Rule::php_code_segment_eof => {
                for inner in pair.into_inner() {
                    let was_interface = matches!(inner.as_rule(), Rule::interface_declaration);
                    let was_trait = matches!(inner.as_rule(), Rule::trait_declaration);
                    if let Some(stmt) = walk_statement(inner)? {
                        if was_interface {
                            if let StmtKind::ClassDecl { ref name, .. } = stmt.kind {
                                interface_names.insert(name.clone());
                            }
                        }
                        if was_trait {
                            if let StmtKind::ClassDecl { ref name, .. } = stmt.kind {
                                trait_names.insert(name.clone());
                            }
                        }
                        body.push(stmt);
                    }
                }
            }
            Rule::php_echo_segment_closed | Rule::php_echo_segment_eof => {
                let expr_pair = pair
                    .into_inner()
                    .next()
                    .ok_or("php_echo_segment missing expression")?;
                let expr = walk_expression(expr_pair)?;
                body.push(Statement::new(StmtKind::Echo(vec![expr])));
            }
            _ => {
                let was_interface = matches!(pair.as_rule(), Rule::interface_declaration);
                let was_trait = matches!(pair.as_rule(), Rule::trait_declaration);
                if let Some(stmt) = walk_statement(pair)? {
                    if was_interface {
                        if let StmtKind::ClassDecl { ref name, .. } = stmt.kind {
                            interface_names.insert(name.clone());
                        }
                    }
                    if was_trait {
                        if let StmtKind::ClassDecl { ref name, .. } = stmt.kind {
                            trait_names.insert(name.clone());
                        }
                    }
                    body.push(stmt);
                }
            }
        }
    }

    Ok(())
}

pub fn parse(source: &str) -> Result<Module, String> {
    let trimmed = source.trim_start();
    let should_normalize_first =
        trimmed.starts_with("<?") || (trimmed.starts_with('<') && source.contains("<?"));
    let mut normalized_source = None;
    let mut pairs = if should_normalize_first {
        let normalized = normalize_mixed_php_source(source);
        if std::env::var_os("VYBEX_DEBUG_WRITE_NORMALIZED_PHP").is_some() {
            let _ = std::fs::write("/tmp/vybex_normalized.php", &normalized);
        }
        normalized_source = Some(normalized);
        PhpParser::parse(
            Rule::program_pure,
            normalized_source.as_deref().unwrap_or(source),
        )
        .map_err(|e| format!("PHP parse error: {}", e))?
    } else {
        match PhpParser::parse(Rule::program_pure, source) {
            Ok(pairs) => pairs,
            Err(_) => PhpParser::parse(Rule::program, source)
                .map_err(|e| format!("PHP parse error: {}", e))?,
        }
    };
    let source = normalized_source.as_deref().unwrap_or(source);

    let mut body = Vec::new();
    let mut interface_names: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut trait_names: std::collections::HashSet<String> = std::collections::HashSet::new();

    // Reset the per-parse trait-usage maps so prior `parse()` calls
    // don't leak state. CLASS_STACK should already be empty here (push
    // and pop are paired inside walk_class_decl/walk_trait_decl/etc).
    TRAIT_USAGES.with(|t| t.borrow_mut().clear());
    TRAIT_ALIASES.with(|t| t.borrow_mut().clear());
    LINE_STARTS.with(|starts| {
        *starts.borrow_mut() = build_line_starts(source);
    });

    let program = pairs.next().ok_or("empty parse")?;
    match program.as_rule() {
        Rule::program => {
            collect_program_body(program, &mut body, &mut interface_names, &mut trait_names)?;
        }
        Rule::program_pure => {
            for pair in program.into_inner() {
                if matches!(pair.as_rule(), Rule::EOI) {
                    continue;
                }
                let pair = if matches!(pair.as_rule(), Rule::pure_top_level_statement) {
                    match pair.into_inner().next() {
                        Some(inner) => inner,
                        None => continue,
                    }
                } else {
                    pair
                };
                let was_interface = matches!(pair.as_rule(), Rule::interface_declaration);
                let was_trait = matches!(pair.as_rule(), Rule::trait_declaration);
                if let Some(stmt) = walk_statement(pair)? {
                    if was_interface {
                        if let StmtKind::ClassDecl { ref name, .. } = stmt.kind {
                            interface_names.insert(name.clone());
                        }
                    }
                    if was_trait {
                        if let StmtKind::ClassDecl { ref name, .. } = stmt.kind {
                            trait_names.insert(name.clone());
                        }
                    }
                    body.push(stmt);
                }
            }
        }
        _ => return Err("unexpected PHP parse root".to_string()),
    }

    // Build a registry of interface const members (interface_name →
    // [const_member_clones]). Walk the body once, find each ClassDecl
    // whose name is in `interface_names`, copy out its Const members.
    let mut interface_consts: std::collections::HashMap<String, Vec<ClassMember>> =
        std::collections::HashMap::new();
    for stmt in &body {
        if let StmtKind::ClassDecl { name, members, .. } = &stmt.kind {
            if interface_names.contains(name) {
                let consts: Vec<ClassMember> = members
                    .iter()
                    .filter(|m| matches!(m, ClassMember::Const { .. }))
                    .cloned()
                    .collect();
                if !consts.is_empty() {
                    interface_consts.insert(name.clone(), consts);
                }
            }
        }
    }

    // Fold interface consts into every class that `implements` them.
    // Skip if the class already declares a const of the same name
    // (PHP shadowing rules). Apply to ClassDecl entries only — the
    // interface entries themselves stay untouched.
    if !interface_consts.is_empty() {
        for stmt in &mut body {
            if let StmtKind::ClassDecl {
                name,
                interfaces,
                members,
                ..
            } = &mut stmt.kind
            {
                if interface_names.contains(name) {
                    continue;
                }
                let existing_const_names: std::collections::HashSet<String> = members
                    .iter()
                    .filter_map(|m| {
                        if let ClassMember::Const { name, .. } = m {
                            Some(name.clone())
                        } else {
                            None
                        }
                    })
                    .collect();
                for iface in interfaces.iter() {
                    if let Some(iface_consts) = interface_consts.get(iface) {
                        for c in iface_consts {
                            if let ClassMember::Const { name: cn, .. } = c {
                                if !existing_const_names.contains(cn) {
                                    members.push(c.clone());
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Build a registry of trait members (name → all members). Traits
    // are walked as ClassDecl by walk_trait_decl; their bodies become
    // available here. Copy const + method members into using classes.
    let mut trait_members: std::collections::HashMap<String, Vec<ClassMember>> =
        std::collections::HashMap::new();
    for stmt in &body {
        if let StmtKind::ClassDecl { name, members, .. } = &stmt.kind {
            if trait_names.contains(name) {
                trait_members.insert(name.clone(), members.clone());
            }
        }
    }

    // Snapshot trait usage map, then fold trait members into using
    // classes. Skip member names already declared on the class (PHP
    // trait conflict rule: class > trait). For class-vs-class duplicates
    // across multiple traits, keep the first one (last-wins would
    // hit the `insteadof` semantic edge cases anyway).
    let usages: std::collections::HashMap<String, Vec<String>> =
        TRAIT_USAGES.with(|t| t.borrow().clone());
    let aliases: std::collections::HashMap<String, Vec<(String, String, String)>> =
        TRAIT_ALIASES.with(|t| t.borrow().clone());
    if !trait_members.is_empty() && !usages.is_empty() {
        for stmt in &mut body {
            if let StmtKind::ClassDecl { name, members, .. } = &mut stmt.kind {
                if trait_names.contains(name) {
                    continue;
                }
                let Some(used) = usages.get(name) else {
                    continue;
                };
                let mut declared: std::collections::HashSet<String> = members
                    .iter()
                    .filter_map(|m| match m {
                        ClassMember::Const { name, .. } => Some(name.clone()),
                        ClassMember::Property { name, .. } => Some(name.clone()),
                        ClassMember::Method(stmt) => {
                            if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                                Some(name.clone())
                            } else {
                                None
                            }
                        }
                        _ => None,
                    })
                    .collect();
                let class_aliases: &[(String, String, String)] =
                    aliases.get(name).map(Vec::as_slice).unwrap_or(&[]);
                for tname in used {
                    if let Some(tmembers) = trait_members.get(tname) {
                        for m in tmembers {
                            let mname = match m {
                                ClassMember::Const { name, .. } => Some(name.clone()),
                                ClassMember::Property { name, .. } => Some(name.clone()),
                                ClassMember::Method(stmt) => {
                                    if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                                        Some(name.clone())
                                    } else {
                                        None
                                    }
                                }
                                _ => None,
                            };
                            if let Some(mn) = mname {
                                if !declared.contains(&mn) {
                                    members.push(m.clone());
                                    declared.insert(mn.clone());
                                }
                                // Apply any alias targeting this trait+method.
                                // The alias triple is (source_trait, method, alias);
                                // an empty source_trait means unqualified
                                // `method as alias` (matches any trait). When
                                // qualified, only apply if THIS trait matches
                                // — that's how `Y::speak as ySpeak` only
                                // creates the ySpeak alias for Y's speak,
                                // not X's.
                                for (src_trait, src, dst) in class_aliases {
                                    if src != &mn {
                                        continue;
                                    }
                                    if !src_trait.is_empty() && src_trait != tname {
                                        continue;
                                    }
                                    if declared.contains(dst) {
                                        continue;
                                    }
                                    if let ClassMember::Method(stmt) = m {
                                        if let StmtKind::FunctionDecl {
                                            params,
                                            return_type,
                                            body: mbody,
                                            modifiers,
                                            handles,
                                            is_async,
                                            is_generator,
                                            is_sub,
                                            ..
                                        } = &stmt.kind
                                        {
                                            let aliased = Statement::new(StmtKind::FunctionDecl {
                                                name: dst.clone(),
                                                params: params.clone(),
                                                return_type: return_type.clone(),
                                                body: mbody.clone(),
                                                modifiers: modifiers.clone(),
                                                handles: handles.clone(),
                                                is_async: *is_async,
                                                is_generator: *is_generator,
                                                is_sub: *is_sub,
                                            });
                                            members.push(ClassMember::Method(Box::new(aliased)));
                                            declared.insert(dst.clone());
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // PHP function/class hoisting: top-level `function foo()` and
    // `class Foo` declarations are visible anywhere in the file, so forward
    // calls like `print hijriDate()` above a later `function hijriDate()`
    // definition are legal. Reorder the body so decls come first — same
    // pattern the JS walker uses.
    let mut hoisted = Vec::new();
    let mut rest = Vec::new();
    for stmt in body {
        if matches!(
            stmt.kind,
            StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. }
        ) {
            hoisted.push(stmt);
        } else {
            rest.push(stmt);
        }
    }
    hoisted.append(&mut rest);
    let body = hoisted;

    LINE_STARTS.with(|starts| starts.borrow_mut().clear());

    Ok(Module {
        name: String::new(),
        language: Lang::PHP,
        body,
        imports: Vec::new(),
    })
}

// ─── Statements ────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let rule = pair.as_rule();
    let kind = match rule {
        Rule::EOI => return Ok(None),
        Rule::function_declaration => walk_function_decl(pair)?,
        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::interface_declaration => walk_interface_decl(pair)?,
        Rule::trait_declaration => walk_trait_decl(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,
        _ => walk_statement_kind(pair, rule)?,
    };
    Ok(Some(Statement::with_span(kind, span)))
}

fn walk_statement_kind(pair: Pair<Rule>, rule: Rule) -> Result<StmtKind, String> {
    let kind = match rule {
        Rule::empty_statement => StmtKind::Empty,

        Rule::block_statement => {
            let inner = pair.into_inner();
            let mut stmts = Vec::new();
            for s in inner {
                if let Some(st) = walk_statement(s)? {
                    stmts.push(st);
                }
            }
            StmtKind::Block(stmts)
        }

        Rule::echo_statement | Rule::print_statement => {
            let exprs: Result<Vec<Expression>, String> = pair
                .into_inner()
                .filter(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .collect();
            StmtKind::Echo(exprs?)
        }

        Rule::expression_statement => {
            let expr = walk_expression(pair.into_inner().next().unwrap())?;
            StmtKind::Expr(expr)
        }

        Rule::const_statement => {
            // const NAME = expr;
            let mut inner = inner_nokw(pair);
            let name = inner.next().unwrap().as_str().to_string();
            let value = walk_expression(inner.next().unwrap())?;
            StmtKind::VarDecl {
                kind: VarDeclKind::Const,
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(value),
                    array_bounds: None,
                    with_events: false,
                }],
            }
        }

        Rule::global_statement => {
            // global $a, $b;  → ScopeDecl { Global, names }
            let names: Vec<String> = pair
                .into_inner()
                .filter(|p| matches!(p.as_rule(), Rule::variable))
                .map(|p| strip_dollar(p.as_str()).to_string())
                .collect();
            StmtKind::ScopeDecl {
                kind: ScopeDeclKind::Global,
                names,
            }
        }

        Rule::template_break_stmt | Rule::template_echo_stmt => {
            // `?>HTML<?php` (or `<?=`) inside a statement list. Emit the
            // literal HTML as `echo "…";`. The trailing `<?= expr ?>` is
            // handled when the enclosing segment reaches the echo block.
            let mut text = String::new();
            for p in pair.into_inner() {
                if matches!(p.as_rule(), Rule::template_text) {
                    text = p.as_str().to_string();
                }
            }
            StmtKind::Echo(vec![Expression::new(ExprKind::Lit(Literal::Str(text)))])
        }

        Rule::static_variable_statement => {
            // `static $x;` or `static $x = expr;` — function-local static
            // variable. We don't yet preserve state across calls (that
            // needs runtime support); compile as a regular VarDecl so
            // subsequent `$x` references resolve as a local and the
            // optional initializer runs on first compile.
            let mut decls = Vec::new();
            for p in pair.into_inner() {
                if matches!(p.as_rule(), Rule::static_variable_decl) {
                    let mut name = String::new();
                    let mut init: Option<Expression> = None;
                    for inner in p.into_inner() {
                        match inner.as_rule() {
                            Rule::variable => name = strip_dollar(inner.as_str()).to_string(),
                            _ => {
                                init = Some(walk_expression(inner)?);
                            }
                        }
                    }
                    decls.push(crate::ast::VarDeclarator {
                        pattern: crate::ast::BindingPattern::Ident(name),
                        init,
                        type_hint: None,
                        array_bounds: None,
                        with_events: false,
                    });
                }
            }
            StmtKind::VarDecl {
                declarations: decls,
                kind: crate::ast::VarDeclKind::Static,
            }
        }

        Rule::namespace_statement => {
            // namespace Foo\Bar; or namespace Foo\Bar { ... }
            // We honour the form but flatten the body — PHP namespace
            // resolution is otherwise cosmetic for our compilation.
            let mut name = String::new();
            let mut body = Vec::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::qualified_name => name = p.as_str().to_string(),
                    Rule::block_statement => {
                        for s in p.into_inner() {
                            if let Some(st) = walk_statement(s)? {
                                body.push(st);
                            }
                        }
                    }
                    _ => {}
                }
            }
            // For the bare `namespace Foo;` form, just discard — return Empty.
            if body.is_empty() {
                StmtKind::Empty
            } else {
                StmtKind::NamespaceDecl { name, body }
            }
        }

        Rule::use_statement => {
            // `use Foo\Bar;` / `use function Foo\bar;` — discard.
            // PHP `use` is for namespace resolution, which we flatten.
            StmtKind::Empty
        }

        Rule::if_statement => walk_if(pair)?,

        Rule::while_statement => {
            let mut inner = inner_nokw(pair);
            let cond = walk_expression(inner.next().unwrap())?;
            let body = walk_statement_into_body(inner.next().unwrap())?;
            StmtKind::While {
                cond,
                body,
                else_body: None,
            }
        }

        Rule::do_while_statement => {
            let mut inner = inner_nokw(pair);
            let body = walk_statement_into_body(inner.next().unwrap())?;
            let cond = walk_expression(inner.next().unwrap())?;
            StmtKind::DoWhile {
                body,
                cond,
                until: false,
            }
        }

        Rule::for_statement => walk_for(pair)?,

        Rule::foreach_statement => walk_foreach(pair)?,

        Rule::goto_statement => {
            let mut label = String::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::identifier => label = p.as_str().to_string(),
                    Rule::expression => {
                        label = match walk_expression(p)?.kind {
                            ExprKind::Ident(name) => name,
                            ExprKind::Lit(Literal::Str(text)) => text,
                            _ => "__php_dynamic_goto".to_string(),
                        };
                    }
                    _ => {}
                }
            }
            StmtKind::GoTo(label)
        }

        Rule::label_statement => {
            let mut inner = pair.into_inner();
            let label = inner.next().unwrap().as_str().to_string();
            if let Some(body_pair) = inner.next() {
                let body =
                    walk_statement(body_pair)?.unwrap_or_else(|| Statement::new(StmtKind::Empty));
                StmtKind::Labeled {
                    label,
                    body: Box::new(body),
                }
            } else {
                StmtKind::Label(label)
            }
        }

        Rule::declare_statement => {
            let body = pair.into_inner().find_map(|p| match p.as_rule() {
                Rule::statement | Rule::block_statement => Some(p),
                _ => None,
            });
            match body {
                Some(stmt) => StmtKind::Block(walk_statement_into_body(stmt)?),
                None => StmtKind::Empty,
            }
        }

        Rule::switch_statement => walk_switch(pair)?,

        Rule::return_statement => {
            let expr = pair
                .into_inner()
                .find(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .transpose()?;
            StmtKind::Return(expr)
        }

        Rule::break_statement => {
            let level = pair
                .into_inner()
                .find(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .transpose()?;
            // PHP break/continue can take an integer level
            let target = match level {
                Some(Expression {
                    kind: ExprKind::Lit(Literal::Int(n)),
                    ..
                }) => BreakTarget::Level(n as u32),
                Some(_) => BreakTarget::Implicit,
                None => BreakTarget::Implicit,
            };
            StmtKind::Break(target)
        }

        Rule::continue_statement => {
            let level = pair
                .into_inner()
                .find(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .transpose()?;
            let target = match level {
                Some(Expression {
                    kind: ExprKind::Lit(Literal::Int(n)),
                    ..
                }) => ContinueTarget::Level(n as u32),
                _ => ContinueTarget::Implicit,
            };
            StmtKind::Continue(target)
        }

        Rule::throw_statement => {
            let expr = walk_expression(inner_nokw(pair).next().unwrap())?;
            StmtKind::Throw {
                expr: Some(expr),
                cause: None,
            }
        }

        Rule::try_statement => walk_try(pair)?,

        other => return Err(format!("walker: unhandled statement rule {:?}", other)),
    };

    Ok(kind)
}

/// Walk a `statement` rule into a `Vec<Statement>` (a body). If the
/// rule is a block, return its contents; otherwise wrap the single
/// statement in a one-element Vec.
fn walk_statement_into_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    if matches!(pair.as_rule(), Rule::block_statement) {
        let mut stmts = Vec::new();
        for s in pair.into_inner() {
            if let Some(st) = walk_statement(s)? {
                stmts.push(st);
            }
        }
        Ok(stmts)
    } else {
        match walk_statement(pair)? {
            Some(s) => Ok(vec![s]),
            None => Ok(Vec::new()),
        }
    }
}

// ─── Control flow ──────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = inner_nokw(pair);
    let cond = walk_expression(inner.next().unwrap())?;
    let then_pair = inner.next().unwrap();
    let (then_body, elifs, else_body) = match then_pair.as_rule() {
        Rule::if_alt_block => walk_if_alt_block(then_pair)?,
        _ => {
            let then_body = walk_statement_into_body(then_pair)?;
            let mut elifs = Vec::new();
            let mut else_body = None;
            for p in inner {
                match p.as_rule() {
                    Rule::elseif_clause => {
                        let mut e = inner_nokw(p);
                        let c = walk_expression(e.next().unwrap())?;
                        let b = walk_statement_into_body(e.next().unwrap())?;
                        elifs.push((c, b));
                    }
                    Rule::else_clause => {
                        let s = inner_nokw(p).next().unwrap();
                        else_body = Some(walk_statement_into_body(s)?);
                    }
                    _ => {}
                }
            }
            (then_body, elifs, else_body)
        }
    };
    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    })
}

fn walk_if_alt_block(
    pair: Pair<Rule>,
) -> Result<
    (
        Vec<Statement>,
        Vec<(Expression, Vec<Statement>)>,
        Option<Vec<Statement>>,
    ),
    String,
> {
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body = None;

    for part in pair.into_inner() {
        match part.as_rule() {
            Rule::elseif_alt_clause => {
                let mut cond = None;
                let mut body = Vec::new();
                for inner in inner_nokw(part) {
                    match inner.as_rule() {
                        Rule::expression => cond = Some(walk_expression(inner)?),
                        _ => {
                            if let Some(stmt) = walk_statement(inner)? {
                                body.push(stmt);
                            }
                        }
                    }
                }
                elifs.push((cond.ok_or("elseif alt block: missing condition")?, body));
            }
            Rule::else_alt_clause => {
                let mut body = Vec::new();
                for inner in inner_nokw(part) {
                    if let Some(stmt) = walk_statement(inner)? {
                        body.push(stmt);
                    }
                }
                else_body = Some(body);
            }
            Rule::kw_endif => {}
            _ => {
                if let Some(stmt) = walk_statement(part)? {
                    then_body.push(stmt);
                }
            }
        }
    }

    Ok((then_body, elifs, else_body))
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // for_statement = { kw_for ~ "(" ~ for_init? ~ ";" ~ expression? ~ ";"
    //                   ~ for_update? ~ ")" ~ statement }
    let mut init: Option<Vec<Expression>> = None;
    let mut cond: Option<Expression> = None;
    let mut update: Option<Vec<Expression>> = None;
    let mut body_stmt: Option<Pair<Rule>> = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::for_init => {
                let exprs: Result<Vec<_>, _> = p.into_inner().map(walk_expression).collect();
                init = Some(exprs?);
            }
            Rule::expression => {
                cond = Some(walk_expression(p)?);
            }
            Rule::for_update => {
                let exprs: Result<Vec<_>, _> = p.into_inner().map(walk_expression).collect();
                update = Some(exprs?);
            }
            _ => {
                body_stmt = Some(p);
            }
        }
    }
    let body = walk_statement_into_body(body_stmt.ok_or("for: missing body")?)?;

    // Compose multi-init / multi-update into a `Sequence` expression
    // wrapped in an Expr statement, since the common AST's `For` only
    // takes a single init Box<Statement> and a single update Expression.
    let init_stmt = init.map(|exprs| {
        let stmt_kind = if exprs.len() == 1 {
            StmtKind::Expr(exprs.into_iter().next().unwrap())
        } else {
            StmtKind::Expr(Expression::new(ExprKind::Sequence(exprs)))
        };
        Box::new(Statement::new(stmt_kind))
    });
    let update_expr = update.map(|exprs| {
        if exprs.len() == 1 {
            exprs.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Sequence(exprs))
        }
    });

    Ok(StmtKind::For {
        init: init_stmt,
        cond,
        update: update_expr,
        body,
    })
}

fn walk_foreach(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // foreach_statement = { kw_foreach ~ "(" ~ expression ~ kw_as
    //                       ~ foreach_target ~ ")" ~ (statement | foreach_alt_block) }
    let mut iter: Option<Expression> = None;
    let mut target_pair: Option<Pair<Rule>> = None;
    let mut body: Vec<Statement> = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression => iter = Some(walk_expression(p)?),
            Rule::foreach_target => target_pair = Some(p),
            Rule::foreach_alt_block => {
                for stmt in p.into_inner() {
                    if matches!(stmt.as_rule(), Rule::kw_endforeach) {
                        continue;
                    }
                    if let Some(walked) = walk_statement(stmt)? {
                        body.push(walked);
                    }
                }
            }
            _ => {
                if matches!(
                    p.as_rule(),
                    Rule::block_statement
                        | Rule::expression_statement
                        | Rule::if_statement
                        | Rule::while_statement
                        | Rule::do_while_statement
                        | Rule::for_statement
                        | Rule::foreach_statement
                        | Rule::switch_statement
                        | Rule::return_statement
                        | Rule::break_statement
                        | Rule::continue_statement
                        | Rule::throw_statement
                        | Rule::try_statement
                        | Rule::echo_statement
                        | Rule::print_statement
                        | Rule::empty_statement
                        | Rule::function_declaration
                        | Rule::class_declaration
                        | Rule::template_break_stmt
                        | Rule::template_echo_stmt
                        | Rule::goto_statement
                        | Rule::label_statement
                        | Rule::declare_statement
                ) {
                    if let Some(walked) = walk_statement(p)? {
                        body.push(walked);
                    }
                }
            }
        }
    }

    let target = target_pair.ok_or("foreach: missing target")?;
    let target_suffix = target.as_span().start();
    // foreach_target = { variable "=>" value-target | value-target }
    let mut tparts = target.into_inner();
    let first = tparts.next().ok_or("foreach: empty target")?;
    let second = tparts.next();

    let (key, value_target) = if let Some(second_var) = second {
        // key => value form
        let k = strip_dollar(first.as_str()).to_string();
        (Some(k), walk_foreach_value_target(second_var)?)
    } else {
        (None, walk_foreach_value_target(first)?)
    };

    if body.is_empty() {
        return Err("foreach: missing body".into());
    }
    let (var, prefix) = foreach_binding_target(value_target, target_suffix)?;
    if let Some(prefix_stmt) = prefix {
        body.insert(0, prefix_stmt);
    }
    Ok(StmtKind::ForIn {
        var,
        key,
        iter: iter.ok_or("foreach: missing iterable")?,
        body,
        of: true, // PHP foreach iterates values, like JS for...of
        else_body: None,
        is_async: false,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = inner_nokw(pair);
    let expr = walk_expression(inner.next().unwrap())?;
    let mut cases = Vec::new();
    let mut default: Option<Vec<Statement>> = None;
    for p in inner {
        if !matches!(p.as_rule(), Rule::switch_case) {
            continue;
        }
        // switch_case = { (kw_case ~ expression | kw_default) ~ ":" ~ statement* }
        // After filtering kw_case/kw_default, the remaining children are
        // [expression?] + [statements...]. We detect "default" by checking
        // the source string since both the kw_* tokens are filtered.
        let case_src = p.as_str();
        let is_default = case_src.trim_start().to_lowercase().starts_with("default");
        let mut case_inner = inner_nokw(p);
        let mut case_value: Option<Expression> = None;
        if !is_default {
            if let Some(e) = case_inner.next() {
                if matches!(e.as_rule(), Rule::expression) {
                    case_value = Some(walk_expression(e)?);
                }
            }
        }
        let body: Result<Vec<Statement>, String> = case_inner
            .filter_map(|p| walk_statement(p).transpose())
            .collect();
        let body = body?;
        if is_default {
            default = Some(body);
        } else {
            cases.push(SwitchCase {
                conditions: vec![CaseCondition::Value(
                    case_value.unwrap_or_else(Expression::null),
                )],
                body,
            });
        }
    }
    Ok(StmtKind::Switch {
        expr,
        cases,
        default,
    })
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = inner_nokw(pair);
    let block = inner.next().unwrap();
    let body = walk_statement_into_body(block)?;
    let mut catches = Vec::new();
    let mut finally: Option<Vec<Statement>> = None;
    for p in inner {
        match p.as_rule() {
            Rule::catch_clause => {
                let mut cat = inner_nokw(p);
                let catch_type = cat.next().unwrap();
                // PHP catches `\UnhandledMatchError $e` — qualified
                // names start with `\` for the global namespace. Strip
                // the leading backslash so the type name matches the
                // canonical exception form the throw site produces.
                let types: Vec<String> = catch_type
                    .into_inner()
                    .map(|q| q.as_str().trim_start_matches('\\').to_string())
                    .collect();
                let mut var: Option<String> = None;
                let mut catch_body_pair: Option<Pair<Rule>> = None;
                for sub in cat {
                    match sub.as_rule() {
                        Rule::variable => var = Some(strip_dollar(sub.as_str()).to_string()),
                        Rule::block_statement => catch_body_pair = Some(sub),
                        _ => {}
                    }
                }
                let catch_body =
                    walk_statement_into_body(catch_body_pair.ok_or("catch: missing body")?)?;
                catches.push(CatchClause {
                    types,
                    var_name: var,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
            }
            Rule::finally_clause => {
                let body = walk_statement_into_body(inner_nokw(p).next().unwrap())?;
                finally = Some(body);
            }
            _ => {}
        }
    }
    Ok(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    })
}

// ─── Function & class declarations ────────────────────────────────────────

/// Recursively scan a function body for `yield` / `yield from` expressions.
/// Does NOT descend into nested function/closure/class bodies — those are
/// their own generator scope.
fn body_contains_yield(stmts: &[Statement]) -> bool {
    enum Node<'a> {
        Stmt(&'a Statement),
        Expr(&'a Expression),
        Place(&'a PlaceExpr),
    }

    let mut stack = stmts.iter().map(Node::Stmt).collect::<Vec<_>>();
    while let Some(node) = stack.pop() {
        match node {
            Node::Place(place) => match place {
                PlaceExpr::Ident(_) => {}
                PlaceExpr::Member { object, .. } => stack.push(Node::Expr(object)),
                PlaceExpr::Index { object, index, .. } => {
                    stack.push(Node::Expr(object));
                    stack.push(Node::Expr(index));
                }
                PlaceExpr::Deref(expr) => stack.push(Node::Expr(expr)),
            },
            Node::Expr(expr) => match &expr.kind {
                ExprKind::Yield(_) | ExprKind::YieldFrom(_) => return true,
                ExprKind::Lambda { .. }
                | ExprKind::FunctionExpr(_)
                | ExprKind::ClassExpr { .. }
                | ExprKind::Lit(_)
                | ExprKind::Ident(_)
                | ExprKind::DefaultOf(_)
                | ExprKind::This
                | ExprKind::Super
                | ExprKind::AddressOf(_)
                | ExprKind::Destructure(_) => {}
                ExprKind::RefOf(place) => stack.push(Node::Place(place)),
                ExprKind::Unary { expr, .. }
                | ExprKind::RefLoad(expr)
                | ExprKind::IsType { expr, .. }
                | ExprKind::Cast { expr, .. }
                | ExprKind::TypeOf(expr)
                | ExprKind::Spread(expr)
                | ExprKind::Await(expr)
                | ExprKind::Void(expr)
                | ExprKind::Delete(expr) => stack.push(Node::Expr(expr)),
                ExprKind::Binary { left, right, .. }
                | ExprKind::NullCoalesce { left, right }
                | ExprKind::Assign {
                    target: left,
                    value: right,
                }
                | ExprKind::Walrus {
                    target: left,
                    value: right,
                }
                | ExprKind::Range {
                    start: left,
                    end: right,
                    ..
                } => {
                    stack.push(Node::Expr(left));
                    stack.push(Node::Expr(right));
                }
                ExprKind::StaticAccess { class, member } => {
                    stack.push(Node::Expr(class));
                    stack.push(Node::Expr(member));
                }
                ExprKind::Ternary { cond, then, else_ } => {
                    stack.push(Node::Expr(cond));
                    stack.push(Node::Expr(then));
                    stack.push(Node::Expr(else_));
                }
                ExprKind::Member { object, .. } => stack.push(Node::Expr(object)),
                ExprKind::Index { object, index, .. } => {
                    stack.push(Node::Expr(object));
                    stack.push(Node::Expr(index));
                }
                ExprKind::Call { callee, args, .. } => {
                    stack.push(Node::Expr(callee));
                    stack.extend(args.iter().map(|arg| Node::Expr(&arg.value)));
                }
                ExprKind::New { class, args } => {
                    stack.push(Node::Expr(class));
                    stack.extend(args.iter().map(|arg| Node::Expr(&arg.value)));
                }
                ExprKind::SuperCall { args, .. } => {
                    stack.extend(args.iter().map(|arg| Node::Expr(&arg.value)));
                }
                ExprKind::Array(elems) => {
                    for elem in elems {
                        stack.push(Node::Expr(&elem.value));
                        if let Some(key) = &elem.key {
                            stack.push(Node::Expr(key));
                        }
                    }
                }
                ExprKind::Tuple(exprs) | ExprKind::Set(exprs) | ExprKind::Sequence(exprs) => {
                    stack.extend(exprs.iter().map(Node::Expr));
                }
                ExprKind::Object(props) => {
                    for prop in props {
                        match prop {
                            ObjectProperty::KeyValue { key, value }
                            | ObjectProperty::Computed { key, value } => {
                                stack.push(Node::Expr(key));
                                stack.push(Node::Expr(value));
                            }
                            ObjectProperty::Spread(expr) => stack.push(Node::Expr(expr)),
                            _ => {}
                        }
                    }
                }
                ExprKind::Interpolation(parts) => {
                    for part in parts {
                        match part {
                            InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                                stack.push(Node::Expr(expr));
                            }
                            _ => {}
                        }
                    }
                }
                ExprKind::Match { subject, arms } => {
                    stack.push(Node::Expr(subject));
                    for arm in arms {
                        if let Some(conditions) = &arm.conditions {
                            stack.extend(conditions.iter().map(Node::Expr));
                        }
                        stack.push(Node::Expr(&arm.body));
                    }
                }
                ExprKind::Comprehension {
                    element,
                    generators,
                    ..
                } => {
                    stack.push(Node::Expr(element));
                    for generator in generators {
                        stack.push(Node::Expr(&generator.iter));
                        stack.extend(generator.conditions.iter().map(Node::Expr));
                    }
                }
                ExprKind::Slice { lower, upper, step } => {
                    if let Some(expr) = lower {
                        stack.push(Node::Expr(expr));
                    }
                    if let Some(expr) = upper {
                        stack.push(Node::Expr(expr));
                    }
                    if let Some(expr) = step {
                        stack.push(Node::Expr(expr));
                    }
                }
            },
            Node::Stmt(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => {}
                StmtKind::Expr(expr) => stack.push(Node::Expr(expr)),
                StmtKind::Block(body)
                | StmtKind::With { body, .. }
                | StmtKind::Using { body, .. }
                | StmtKind::Lock { body, .. }
                | StmtKind::NamespaceDecl { body, .. } => {
                    stack.extend(body.iter().map(Node::Stmt));
                }
                StmtKind::VarDecl { declarations, .. } => {
                    for decl in declarations {
                        if let Some(init) = &decl.init {
                            stack.push(Node::Expr(init));
                        }
                    }
                }
                StmtKind::Return(expr) => {
                    if let Some(expr) = expr {
                        stack.push(Node::Expr(expr));
                    }
                }
                StmtKind::If {
                    cond,
                    then_body,
                    elifs,
                    else_body,
                } => {
                    stack.push(Node::Expr(cond));
                    stack.extend(then_body.iter().map(Node::Stmt));
                    for (cond, body) in elifs {
                        stack.push(Node::Expr(cond));
                        stack.extend(body.iter().map(Node::Stmt));
                    }
                    if let Some(body) = else_body {
                        stack.extend(body.iter().map(Node::Stmt));
                    }
                }
                StmtKind::While {
                    cond,
                    body,
                    else_body,
                } => {
                    stack.push(Node::Expr(cond));
                    stack.extend(body.iter().map(Node::Stmt));
                    if let Some(body) = else_body {
                        stack.extend(body.iter().map(Node::Stmt));
                    }
                }
                StmtKind::DoWhile { body, cond, .. } => {
                    stack.extend(body.iter().map(Node::Stmt));
                    stack.push(Node::Expr(cond));
                }
                StmtKind::For {
                    init,
                    cond,
                    update,
                    body,
                } => {
                    if let Some(init) = init {
                        stack.push(Node::Stmt(init));
                    }
                    if let Some(cond) = cond {
                        stack.push(Node::Expr(cond));
                    }
                    if let Some(update) = update {
                        stack.push(Node::Expr(update));
                    }
                    stack.extend(body.iter().map(Node::Stmt));
                }
                StmtKind::ForIn {
                    iter,
                    body,
                    else_body,
                    ..
                } => {
                    stack.push(Node::Expr(iter));
                    stack.extend(body.iter().map(Node::Stmt));
                    if let Some(body) = else_body {
                        stack.extend(body.iter().map(Node::Stmt));
                    }
                }
                StmtKind::Switch {
                    expr,
                    cases,
                    default,
                } => {
                    stack.push(Node::Expr(expr));
                    for case in cases {
                        stack.extend(case.body.iter().map(Node::Stmt));
                    }
                    if let Some(body) = default {
                        stack.extend(body.iter().map(Node::Stmt));
                    }
                }
                StmtKind::Try {
                    body,
                    catches,
                    else_body,
                    finally,
                } => {
                    stack.extend(body.iter().map(Node::Stmt));
                    for catch in catches {
                        stack.extend(catch.body.iter().map(Node::Stmt));
                    }
                    if let Some(body) = else_body {
                        stack.extend(body.iter().map(Node::Stmt));
                    }
                    if let Some(body) = finally {
                        stack.extend(body.iter().map(Node::Stmt));
                    }
                }
                StmtKind::Assign { targets, value } => {
                    stack.extend(targets.iter().map(Node::Expr));
                    stack.push(Node::Expr(value));
                }
                StmtKind::CompoundAssign { target, value, .. } => {
                    stack.push(Node::Expr(target));
                    stack.push(Node::Expr(value));
                }
                StmtKind::Throw { expr, cause } => {
                    if let Some(expr) = expr {
                        stack.push(Node::Expr(expr));
                    }
                    if let Some(cause) = cause {
                        stack.push(Node::Expr(cause));
                    }
                }
                StmtKind::Labeled { body, .. } => stack.push(Node::Stmt(body)),
                StmtKind::Echo(exprs) | StmtKind::Delete(exprs) => {
                    stack.extend(exprs.iter().map(Node::Expr));
                }
                StmtKind::Export {
                    declaration,
                    default,
                    ..
                } => {
                    if let Some(declaration) = declaration {
                        stack.push(Node::Stmt(declaration));
                    }
                    if let Some(default) = default {
                        stack.push(Node::Expr(default));
                    }
                }
                StmtKind::MatchStatement { subject, cases } => {
                    stack.push(Node::Expr(subject));
                    for case in cases {
                        if let Some(guard) = &case.guard {
                            stack.push(Node::Expr(guard));
                        }
                        stack.extend(case.body.iter().map(Node::Stmt));
                    }
                }
                StmtKind::Assert { test, msg } => {
                    stack.push(Node::Expr(test));
                    if let Some(msg) = msg {
                        stack.push(Node::Expr(msg));
                    }
                }
                _ => {}
            },
        }
    }
    false
}

fn walk_function_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut params: Vec<Param> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();
    let mut return_type: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier | Rule::method_ident => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::return_type_annotation => {
                return_type = Some(p.as_str().trim_start_matches(':').trim().to_string());
            }
            Rule::block_statement => {
                body = walk_statement_into_body(p)?;
            }
            _ => {}
        }
    }

    body = lower_php_runtime_arg_helpers_in_block(&mut params, body);

    let is_generator = body_contains_yield(&body);
    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator,
        is_sub: false,
    })
}

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if !matches!(p.as_rule(), Rule::param) {
            continue;
        }
        out.push(walk_param(p)?.0);
    }
    Ok(out)
}

/// Walk a param list, returning both the params AND any promotion
/// modifiers (`public` / `private` / `protected` / `readonly`).
/// Used by `__construct` to synthesize property fields and `$this->X = $X`
/// assignments — see PHP 8 promoted constructor parameters.
fn walk_params_with_promotion(
    pair: Pair<Rule>,
) -> Result<Vec<(Param, Option<Visibility>)>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if !matches!(p.as_rule(), Rule::param) {
            continue;
        }
        out.push(walk_param(p)?);
    }
    Ok(out)
}

fn walk_param(pair: Pair<Rule>) -> Result<(Param, Option<Visibility>), String> {
    let raw = pair.as_str();
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut default: Option<Expression> = None;
    let mut promotion: Option<Visibility> = None;
    let pass_by = if raw.contains('&') {
        PassBy::Ref
    } else {
        PassBy::Value
    };
    let is_rest = raw.contains("...");
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_modifier => {
                // Presence of `public`/`private`/`protected`/`readonly`
                // marks this as a promoted constructor param. Default
                // (when the modifier is `readonly` only) is Public per
                // PHP semantics.
                let kw = p.as_str().to_lowercase();
                let v = match kw.as_str() {
                    "private" => Visibility::Private,
                    "protected" => Visibility::Protected,
                    _ => Visibility::Public,
                };
                promotion = Some(v);
            }
            Rule::type_annotation => type_hint = Some(p.as_str().to_string()),
            Rule::variable => name = strip_dollar(p.as_str()).to_string(),
            Rule::expression => default = Some(walk_expression(p)?),
            _ => {}
        }
    }
    let is_optional = default.is_some();
    Ok((
        Param {
            name,
            type_hint,
            default,
            pass_by,
            is_rest,
            is_kwargs: false,
            is_optional,
            is_nullable: false,
        },
        promotion,
    ))
}

fn walk_class_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // Inspect the source slice to know whether the first qualified_name
    // follows `extends` (parent class) or `implements` (interface
    // list). Pest doesn't yield `kw_extends` / `kw_implements` as
    // child pairs, so we'd otherwise misclassify
    // `class Foo implements Bar` (Bar is the parent) — leading to the
    // compiler instantiating a non-existent Bar at construction time.
    let raw = pair.as_str();
    let has_extends = raw.contains(" extends ");
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut modifiers = ClassModifiers::default();
    let mut first_qualified = true;
    let mut deferred_members: Vec<Pair<Rule>> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifier => {
                let s = p.as_str().to_lowercase();
                match s.as_str() {
                    "abstract" => modifiers.is_abstract = true,
                    "final" => modifiers.is_sealed = true,
                    _ => {}
                }
            }
            Rule::identifier | Rule::method_ident if name.is_empty() => {
                name = p.as_str().to_string()
            }
            Rule::qualified_name => {
                if first_qualified && has_extends {
                    parents.push(p.as_str().to_string());
                } else {
                    interfaces.push(p.as_str().to_string());
                }
                first_qualified = false;
            }
            Rule::use_trait
            | Rule::class_constant
            | Rule::property_declaration
            | Rule::method_declaration
            | Rule::empty_statement => {
                deferred_members.push(p);
            }
            _ => {}
        }
    }

    // Push class context BEFORE walking members so `self::` inside
    // method bodies resolves to this class's name.
    push_class_context(&name);
    let walk_result: Result<(), String> = (|| {
        for p in deferred_members {
            if let Some(member) = walk_class_member(p)? {
                members.push(member);
            }
        }
        Ok(())
    })();
    pop_class_context();
    walk_result?;

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers,
        decorators: vec![],
    })
}

fn walk_interface_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // PHP interfaces behave like classes that carry constants. Walking
    // them as ClassDecl lets `Interface::CONST` static access resolve
    // through the standard class-const path. Method signatures are
    // dropped (no bodies on interface methods) — they're documentation
    // only at the AST level.
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut class_members: Vec<ClassMember> = Vec::new();
    let mut deferred: Vec<Pair<Rule>> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier | Rule::method_ident if name.is_empty() => {
                name = p.as_str().to_string()
            }
            Rule::qualified_name => parents.push(p.as_str().to_string()),
            Rule::class_constant => {
                deferred.push(p);
            }
            Rule::method_declaration => {
                // Skip — interface methods have no body, so nothing to
                // emit. Implementing classes provide their own.
            }
            _ => {}
        }
    }

    push_class_context(&name);
    let walk_result: Result<(), String> = (|| {
        for p in deferred {
            if let Some(m) = walk_class_member(p)? {
                class_members.push(m);
            }
        }
        Ok(())
    })();
    pop_class_context();
    walk_result?;

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members: class_members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_trait_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // Treat a trait as a regular class for compilation purposes —
    // `use TraitName` inside another class flattens via the same parent
    // chain.
    let mut name = String::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut deferred_members: Vec<Pair<Rule>> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier | Rule::method_ident if name.is_empty() => {
                name = p.as_str().to_string()
            }
            Rule::use_trait
            | Rule::class_constant
            | Rule::property_declaration
            | Rule::method_declaration
            | Rule::empty_statement => {
                deferred_members.push(p);
            }
            _ => {}
        }
    }

    push_class_context(&name);
    let walk_result: Result<(), String> = (|| {
        for p in deferred_members {
            if let Some(member) = walk_class_member(p)? {
                members.push(member);
            }
        }
        Ok(())
    })();
    pop_class_context();
    walk_result?;

    Ok(StmtKind::ClassDecl {
        name,
        parents: Vec::new(),
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_enum_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut backing_type: Option<String> = None;
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<EnumMember> = Vec::new();
    let mut body_members: Vec<ClassMember> = Vec::new();
    let mut deferred: Vec<Pair<Rule>> = Vec::new();

    // First pass: extract name + simple metadata (no expression walks yet).
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier | Rule::method_ident if name.is_empty() => {
                name = p.as_str().to_string()
            }
            Rule::identifier => backing_type = Some(p.as_str().to_string()),
            Rule::qualified_name => interfaces.push(p.as_str().to_string()),
            Rule::enum_case | Rule::class_constant | Rule::method_declaration | Rule::use_trait => {
                deferred.push(p);
            }
            _ => {}
        }
    }

    // Second pass: walk member expressions/bodies with class context
    // pushed so `self::` inside enum methods/cases resolves to the enum.
    push_class_context(&name);
    let walk_result: Result<(), String> = (|| {
        for p in deferred {
            match p.as_rule() {
                Rule::enum_case => {
                    let mut case_name = String::new();
                    let mut backing: Option<Expression> = None;
                    for c in p.into_inner() {
                        match c.as_rule() {
                            Rule::identifier => case_name = c.as_str().to_string(),
                            Rule::expression => backing = Some(walk_expression(c)?),
                            _ => {}
                        }
                    }
                    members.push(EnumMember {
                        name: case_name,
                        value: backing,
                        constructor_args: Vec::new(),
                    });
                }
                Rule::class_constant | Rule::method_declaration | Rule::use_trait => {
                    if let Some(m) = walk_class_member(p)? {
                        body_members.push(m);
                    }
                }
                _ => {}
            }
        }
        Ok(())
    })();
    pop_class_context();
    walk_result?;

    // Desugar the PHP enum into a plain class so the shared, language-agnostic
    // compiler handles it with NO `profile.name == "php"` branch. Each case
    // becomes a static singleton field `new EnumName(name, value)`; the
    // `cases()`/`from()`/`tryFrom()` helpers are synthesized as static methods.
    // This mirrors exactly what the former `compile_php_enum_decl` emitted, but
    // expressed as common AST in the walker (where language idioms belong).
    let static_mods = || Modifiers {
        is_static: true,
        ..Modifiers::default()
    };
    let mk_param = |param_name: &str| Param {
        name: param_name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };
    let case_ref = |case_name: &str| {
        Expression::new(ExprKind::StaticAccess {
            class: Box::new(Expression::ident(&name)),
            member: Box::new(Expression::ident(case_name)),
        })
    };

    let mut class_members: Vec<ClassMember> = Vec::new();

    // Synthetic constructor: (name, value) → $this->name / $this->value.
    let assign_this = |field: &str, value: Expression| {
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: field.to_string(),
                null_safe: false,
            })],
            value,
        })
    };
    class_members.push(ClassMember::Constructor {
        params: vec![mk_param("__enum_name"), mk_param("__enum_value")],
        body: vec![
            assign_this("name", Expression::ident("__enum_name")),
            assign_this("value", Expression::ident("__enum_value")),
        ],
        base_args: None,
        initializer_target: crate::ast::ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    });

    // User-declared methods / constants.
    class_members.extend(body_members.iter().cloned());

    // cases(): array of all case singletons.
    if !members.is_empty() {
        let elements: Vec<ArrayElement> = members
            .iter()
            .map(|member| ArrayElement {
                key: None,
                value: case_ref(&member.name),
                spread: false,
                by_ref: false,
            })
            .collect();
        class_members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "cases".to_string(),
                params: vec![],
                return_type: Some("array".to_string()),
                body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                    ExprKind::Array(elements),
                ))))],
                modifiers: static_mods(),
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            },
        ))));
    }

    // from()/tryFrom(): match a backing value to its case.
    if backing_type.is_some() && members.iter().any(|member| member.value.is_some()) {
        let build_match_chain = |fallback: Statement| -> Vec<Statement> {
            let mut body = Vec::new();
            for member in &members {
                if let Some(backing) = member.value.clone() {
                    let cond = Expression::new(ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(Expression::ident("v")),
                        right: Box::new(backing),
                    });
                    body.push(Statement::new(StmtKind::If {
                        cond,
                        then_body: vec![Statement::new(StmtKind::Return(Some(case_ref(
                            &member.name,
                        ))))],
                        elifs: Vec::new(),
                        else_body: None,
                    }));
                }
            }
            body.push(fallback);
            body
        };
        class_members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "tryFrom".to_string(),
                params: vec![mk_param("v")],
                return_type: None,
                body: build_match_chain(Statement::new(StmtKind::Return(Some(Expression::null())))),
                modifiers: static_mods(),
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            },
        ))));
        class_members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "from".to_string(),
                params: vec![mk_param("v")],
                return_type: None,
                body: build_match_chain(Statement::new(StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("Error")),
                        args: vec![Argument::positional(Expression::string(&format!(
                            "Invalid backing value for enum \"{}\"",
                            name
                        )))],
                    })),
                    cause: None,
                })),
                modifiers: static_mods(),
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            },
        ))));
    }

    // Each case → a static singleton field initialised to a fresh instance.
    // No backing value ⇒ `value` defaults to the case name (PHP semantics).
    for member in &members {
        let value_expr = member
            .value
            .clone()
            .unwrap_or_else(|| Expression::string(&member.name));
        let init = Expression::new(ExprKind::New {
            class: Box::new(Expression::ident(&name)),
            args: vec![
                Argument::positional(Expression::string(&member.name)),
                Argument::positional(value_expr),
            ],
        });
        class_members.push(ClassMember::Field {
            name: member.name.clone(),
            type_hint: Some(name.clone()),
            init: Some(init),
            modifiers: static_mods(),
            with_events: false,
            array_bounds: None,
        });
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents: vec![],
        interfaces,
        members: class_members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_class_member(pair: Pair<Rule>) -> Result<Option<ClassMember>, String> {
    match pair.as_rule() {
        Rule::empty_statement => Ok(None),
        Rule::use_trait => {
            // `use TraitName(, OtherTrait)*;` or
            // `use TraitName(, OtherTrait)* { adaptations }` inside a
            // class. Record trait names + alias adaptations against
            // the enclosing class so the post-pass in `parse()` can
            // copy trait const + method members and add aliases on
            // the using class. Returns `None` because the trait usage
            // itself is metadata; actual members get injected later.
            if let Some(class_name) = current_class_name() {
                let mut trait_names: Vec<String> = Vec::new();
                let mut aliases: Vec<(String, String, String)> = Vec::new();
                for p in pair.into_inner() {
                    match p.as_rule() {
                        Rule::qualified_name => trait_names.push(p.as_str().to_string()),
                        Rule::trait_adaptation => {
                            // Two forms:
                            //   trait_method_ref ~ "insteadof" ~ qualified_name+ ~ ";"
                            //   trait_method_ref ~ "as" ~ visibility? ~ method_ident? ~ ";"
                            // We only care about the `as` form here —
                            // `insteadof` is handled implicitly by the
                            // first-trait-wins copy order.
                            let raw = p.as_str();
                            let is_alias = raw.contains(" as ");
                            if !is_alias {
                                continue;
                            }
                            let mut method_trait: String = String::new();
                            let mut method_name: Option<String> = None;
                            let mut alias_name: Option<String> = None;
                            for q in p.into_inner() {
                                match q.as_rule() {
                                    Rule::trait_method_ref => {
                                        // trait_method_ref = qualified_name "::" method_ident | method_ident
                                        let mut tname: Option<String> = None;
                                        let mut last_method: Option<String> = None;
                                        for r in q.into_inner() {
                                            match r.as_rule() {
                                                Rule::qualified_name => {
                                                    tname = Some(r.as_str().to_string())
                                                }
                                                Rule::method_ident => {
                                                    last_method = Some(r.as_str().to_string())
                                                }
                                                _ => {}
                                            }
                                        }
                                        if let Some(t) = tname {
                                            method_trait = t;
                                        }
                                        method_name = last_method;
                                    }
                                    Rule::method_ident => {
                                        alias_name = Some(q.as_str().to_string());
                                    }
                                    _ => {}
                                }
                            }
                            if let (Some(src), Some(dst)) = (method_name, alias_name) {
                                aliases.push((method_trait, src, dst));
                            }
                        }
                        _ => {}
                    }
                }
                if !trait_names.is_empty() {
                    TRAIT_USAGES.with(|t| {
                        t.borrow_mut()
                            .entry(class_name.clone())
                            .or_default()
                            .extend(trait_names);
                    });
                }
                if !aliases.is_empty() {
                    TRAIT_ALIASES.with(|t| {
                        t.borrow_mut()
                            .entry(class_name)
                            .or_default()
                            .extend(aliases);
                    });
                }
            }
            Ok(None)
        }
        Rule::class_constant => {
            let mut name = String::new();
            let mut value: Option<Expression> = None;
            let mut visibility = Visibility::Public;
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::member_modifier => visibility = parse_visibility(p.as_str(), visibility),
                    Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
                    Rule::expression => value = Some(walk_expression(p)?),
                    _ => {}
                }
            }
            Ok(Some(ClassMember::Const {
                name,
                type_hint: None,
                value: value.unwrap_or_else(Expression::null),
                visibility,
            }))
        }
        Rule::property_declaration => {
            let mut name = String::new();
            let mut type_hint: Option<String> = None;
            let mut init: Option<Expression> = None;
            let mut modifiers = Modifiers::default();
            let mut getter: Option<Vec<Statement>> = None;
            let mut setter: Option<PropertySetter> = None;
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::member_modifier | Rule::asymmetric_set_modifier => {
                        apply_member_modifier(&mut modifiers, p.as_str())
                    }
                    Rule::type_annotation => type_hint = Some(p.as_str().to_string()),
                    Rule::simple_variable | Rule::variable => {
                        // Property names are stored WITHOUT the `$`
                        // sigil (member access `$this->prop` looks up
                        // "prop"). PHP variables in expression context
                        // keep the `$` (separate namespace), but
                        // property declarations strip it here.
                        let raw = p.as_str();
                        name = raw.strip_prefix('$').unwrap_or(raw).to_string();
                    }
                    Rule::expression => init = Some(walk_expression(p)?),
                    Rule::property_hook_block => {
                        let (hook_getter, hook_setter) = walk_property_hooks(p, &type_hint)?;
                        getter = hook_getter;
                        setter = hook_setter;
                    }
                    _ => {}
                }
            }
            if getter.is_some() || setter.is_some() {
                Ok(Some(ClassMember::Property {
                    name,
                    type_hint,
                    getter,
                    setter,
                    is_auto: false,
                    modifiers,
                }))
            } else {
                Ok(Some(ClassMember::Field {
                    name,
                    type_hint,
                    init,
                    modifiers,
                    with_events: false,
                    array_bounds: None,
                }))
            }
        }
        Rule::method_declaration => {
            let mut method_name = String::new();
            let mut params: Vec<Param> = Vec::new();
            let mut promoted: Vec<(String, Option<String>, Visibility)> = Vec::new();
            let mut body: Vec<Statement> = Vec::new();
            let mut return_type: Option<String> = None;
            let mut modifiers = Modifiers::default();
            let mut has_body = false;
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::member_modifier => apply_member_modifier(&mut modifiers, p.as_str()),
                    Rule::identifier | Rule::method_ident => method_name = p.as_str().to_string(),
                    Rule::param_list => {
                        // Capture both the params and any promotion
                        // visibility so we can synthesize property
                        // assignments below for `__construct`.
                        let with_prom = walk_params_with_promotion(p)?;
                        params = Vec::with_capacity(with_prom.len());
                        for (param, vis) in with_prom {
                            if let Some(v) = vis {
                                promoted.push((param.name.clone(), param.type_hint.clone(), v));
                            }
                            params.push(param);
                        }
                    }
                    Rule::return_type_annotation => {
                        return_type = Some(p.as_str().trim_start_matches(':').trim().to_string());
                    }
                    Rule::block_statement => {
                        body = walk_statement_into_body(p)?;
                        has_body = true;
                    }
                    _ => {}
                }
            }

            // PHP `__construct` becomes a Constructor class member so
            // the compiler-side child-class flow recognises it.
            if method_name == "__construct" {
                // PHP 8 promoted constructor params: each
                // `public/private/protected/readonly $foo` produces
                // (1) a `$this->foo = $foo` assignment prepended to
                // the body, and (2) a property field on the class.
                // The walker emits the assignments here; the property
                // declarations are returned to walk_class_decl through
                // a side channel.
                if !promoted.is_empty() {
                    let mut prelude: Vec<Statement> = Vec::with_capacity(promoted.len());
                    for (pname, _ptype, _pvis) in &promoted {
                        // `pname` retains the `$` sigil per PHP variable
                        // canonicalization (see strip_dollar). The
                        // member name strips the `$` (it's a property,
                        // not a variable); the assignment value uses
                        // the variable form.
                        let field_name = pname.strip_prefix('$').unwrap_or(pname).to_string();
                        let this_expr = Expression::new(ExprKind::This);
                        let target = Expression::new(ExprKind::Member {
                            object: Box::new(this_expr),
                            field: field_name,
                            null_safe: false,
                        });
                        let value = Expression::new(ExprKind::Ident(pname.clone()));
                        let assign = Expression::new(ExprKind::Assign {
                            target: Box::new(target),
                            value: Box::new(value),
                        });
                        prelude.push(Statement::new(StmtKind::Expr(assign)));
                    }
                    prelude.extend(body.drain(..));
                    body = prelude;
                }
                let _ = (return_type, has_body);
                return Ok(Some(ClassMember::Constructor {
                    params,
                    body,
                    base_args: None,
                    initializer_target: crate::ast::ConstructorInitializerTarget::Base,
                    visibility: modifiers.visibility,
                }));
            }

            // Build a Method wrapping a FunctionDecl.
            let method_body = if has_body { body } else { Vec::new() };
            let is_generator = body_contains_yield(&method_body);
            let stmt = Statement::new(StmtKind::FunctionDecl {
                name: method_name,
                params,
                return_type,
                body: method_body,
                modifiers,
                handles: Vec::new(),
                is_async: false,
                is_generator,
                is_sub: false,
            });
            Ok(Some(ClassMember::Method(Box::new(stmt))))
        }
        _ => Ok(None),
    }
}

fn apply_member_modifier(mods: &mut Modifiers, kw: &str) {
    let lower = kw.to_lowercase();
    if lower.starts_with("public") {
        mods.visibility = Visibility::Public;
    } else if lower.starts_with("private") {
        mods.visibility = Visibility::Private;
    } else if lower.starts_with("protected") {
        mods.visibility = Visibility::Protected;
    } else {
        match lower.as_str() {
            "static" => mods.is_static = true,
            "abstract" => mods.is_abstract = true,
            "final" => mods.is_not_overridable = true,
            "readonly" => mods.is_readonly = true,
            _ => {}
        }
    }
}

fn parse_visibility(s: &str, default: Visibility) -> Visibility {
    match s.to_lowercase().as_str() {
        "public" => Visibility::Public,
        "private" => Visibility::Private,
        "protected" => Visibility::Protected,
        _ => default,
    }
}

// ─── Expressions ──────────────────────────────────────────────────────────

fn single_child<'i>(pair: &Pair<'i, Rule>) -> Option<Pair<'i, Rule>> {
    let mut inner = pair.clone().into_inner();
    let first = inner.next()?;
    if inner.next().is_none() {
        Some(first)
    } else {
        None
    }
}

fn transparent_expression_child<'i>(pair: &Pair<'i, Rule>) -> Option<Pair<'i, Rule>> {
    match pair.as_rule() {
        Rule::expression
        | Rule::assignment_expression
        | Rule::logical_or_expression
        | Rule::ternary_expression
        | Rule::null_coalesce_expression
        | Rule::logic_or_expression
        | Rule::logic_and_expression
        | Rule::bit_or_expression
        | Rule::bit_xor_expression
        | Rule::bit_and_expression
        | Rule::equality_expression
        | Rule::comparison_expression
        | Rule::shift_expression
        | Rule::additive_expression
        | Rule::multiplicative_expression
        | Rule::unary_expression
        | Rule::postfix_expression
        | Rule::primary_expression
        | Rule::parenthesized_expression
        | Rule::include_operand => single_child(pair),
        _ => None,
    }
}

fn walk_expression(mut pair: Pair<Rule>) -> Result<Expression, String> {
    while let Some(child) = transparent_expression_child(&pair) {
        pair = child;
    }

    let span = to_span(&pair);
    let rule = pair.as_rule();
    let kind = match rule {
        Rule::expression => return walk_expression(pair.into_inner().next().unwrap()),
        Rule::assignment_expression => return walk_assignment(pair),
        Rule::yield_expression => return walk_yield(pair),
        Rule::null_coalesce_expression => {
            // Right-associative `??` — pest doesn't yield the literal
            // `??` as a child pair (no named operator rule), so the
            // generic walk_left_assoc_binary path can't recover the
            // operator name. Walk the children directly: first is the
            // left operand, optional second is the recursive right.
            let span = to_span(&pair);
            let mut inner = pair.into_inner();
            let left = walk_expression(inner.next().unwrap())?;
            return match inner.next() {
                Some(rhs_pair) => {
                    let right = walk_expression(rhs_pair)?;
                    Ok(Expression::with_span(
                        ExprKind::NullCoalesce {
                            left: Box::new(left),
                            right: Box::new(right),
                        },
                        span,
                    ))
                }
                None => Ok(left),
            };
        }
        Rule::logical_or_expression
        | Rule::logic_or_expression
        | Rule::logic_and_expression
        | Rule::bit_or_expression
        | Rule::bit_xor_expression
        | Rule::bit_and_expression
        | Rule::equality_expression
        | Rule::comparison_expression
        | Rule::shift_expression
        | Rule::additive_expression
        | Rule::multiplicative_expression => return walk_left_assoc_binary(pair),
        Rule::ternary_expression => return walk_ternary(pair),
        Rule::unary_expression => return walk_unary(pair),
        Rule::cast_expression => return walk_cast(pair),
        Rule::postfix_expression => return walk_postfix(pair),

        Rule::primary_expression => return walk_expression(pair.into_inner().next().unwrap()),
        Rule::parenthesized_expression => {
            return walk_expression(pair.into_inner().next().unwrap());
        }

        Rule::literal => return walk_literal(pair),
        Rule::number_lit => return Ok(Expression::with_span(walk_number(&pair).kind, span)),
        Rule::string_lit => return Ok(Expression::with_span(walk_string(&pair).kind, span)),

        Rule::variable => return walk_php_variable_expr(pair, span),
        Rule::simple_variable => ExprKind::Ident(strip_dollar(pair.as_str()).to_string()),
        Rule::identifier | Rule::method_ident => {
            let name = pair.as_str();
            // Bare PHP global constants get rewritten to their JS-shaped
            // equivalent (`M_PI` → `Math.PI`, `STR_PAD_LEFT` → `0`, etc.)
            // so the rest of the pipeline never sees a PHP-specific name.
            if let Some(kind) = php_constant_expr(name, &span) {
                kind
            } else {
                ExprKind::Ident(name.to_string())
            }
        }
        Rule::qualified_name => {
            // Preserve the qualified path so the compiler can resolve it
            // against the profile's `host_packages` list. A qualified
            // name whose first segment matches a host package (e.g.
            // `\Vybe\Http\Response\set_status`) becomes a Component
            // Model host call at compile time. Anything else still
            // resolves by last-segment (user namespaces are flattened
            // for now — worth revisiting when we model user namespaces).
            //
            // Bare PHP constants (M_PI, STR_PAD_LEFT, ...) reach this
            // arm via the qualified_name grammar rule when the parser
            // routes them through the namespace-resolution path. They
            // are rewritten to JS-shaped AST exactly like the
            // identifier arm above — same code path downstream.
            let s = pair.as_str().trim_start_matches('\\');
            if !s.contains('\\') {
                if let Some(kind) = php_constant_expr(s, &span) {
                    kind
                } else {
                    ExprKind::Ident(s.to_string())
                }
            } else {
                ExprKind::Ident(s.to_string())
            }
        }

        Rule::kw_self => {
            // PHP `self::X` is a compile-time reference to the enclosing
            // class. When inside a class member (walk_class_decl /
            // walk_trait_decl / walk_enum_decl pushed the class name onto
            // CLASS_STACK), rewrite `self` to the class name directly so
            // `self::CONST` becomes `ClassName::CONST` and resolves via
            // the static field on the class object. Without context (rare
            // — `self` outside a class is illegal PHP) fall back to
            // `This` so existing call-site code keeps working.
            match current_class_name() {
                Some(cn) => ExprKind::Ident(cn),
                None => ExprKind::This,
            }
        }
        Rule::kw_parent => ExprKind::Super,
        Rule::kw_static => {
            // PHP `static::X` (late static binding) resolves to the
            // calling class at runtime — same `$this` slot that the
            // static-method dispatch puts the class object into. Walk
            // `static` to `This` so `static::X` becomes
            // `StaticAccess { class: This, member: X }`. The compiler
            // then emits `LOCAL_GET this; STRUCT_GET "X"` — for static
            // method calls dispatched via `Class.method()` (Member
            // shape), the `$this` slot holds the class object, so
            // STRUCT_GET on it returns the class const / static field.
            ExprKind::This
        }

        Rule::new_expression => return walk_new(pair),
        Rule::clone_expression => {
            let inner = inner_nokw(pair).next().unwrap();
            let arg = walk_expression(inner)?;
            // PHP `clone $obj` — produce a shallow copy and invoke
            // the class's `__clone` magic method on the copy if one
            // is defined. The walker calls into the
            // `__php_clone_helper` adapter (Rust opcode emitter in
            // `string_adapter.rs::emit_php_clone`) which handles
            // both the shallow copy + magic-method dispatch.
            ExprKind::Call {
                callee: Box::new(Expression::ident("__php_clone_helper")),
                args: vec![Argument::positional(arg)],
                optional: false,
            }
        }
        Rule::match_expression => return walk_match(pair),
        Rule::isset_expression => {
            // Walk each arg with the LHS-depth flag set so the
            // property-access walker doesn't wrap `$obj->prop` in the
            // magic-`__get` ternary. `isset` needs the raw l-value
            // (or our own `__isset` dispatch wrap) — going through
            // `__get` would coerce undefined to the user's fallback
            // and report wrong existence.
            let mut args: Vec<Argument> = Vec::new();
            for p in pair.into_inner() {
                if matches!(p.as_rule(), Rule::expression) {
                    ASSIGN_LHS_DEPTH.with(|d| *d.borrow_mut() += 1);
                    let walked = walk_expression(p);
                    ASSIGN_LHS_DEPTH.with(|d| {
                        let mut bd = d.borrow_mut();
                        *bd = bd.saturating_sub(1);
                    });
                    let walked = walked?;
                    // PHP `isset($obj->prop)` magic dispatch: if `prop`
                    // isn't set on the instance and the class defines
                    // `__isset`, route through it. Otherwise bool-test
                    // whether the direct read is non-null. Walker only
                    // rewrites simple `$var->prop` Member access — more
                    // complex shapes fall through to the regular
                    // `isset(value)` check.
                    let arg_expr = match &walked.kind {
                        ExprKind::Member {
                            object,
                            field,
                            null_safe,
                        } if !*null_safe
                            && !field.starts_with("__")
                            && !matches!(object.kind, ExprKind::This) =>
                        {
                            let obj = (**object).clone();
                            let field = field.clone();
                            build_magic_isset_rewrite(obj, field, &span)
                        }
                        _ => walked,
                    };
                    args.push(Argument::positional(arg_expr));
                }
            }
            ExprKind::Call {
                callee: Box::new(Expression::ident("isset")),
                args,
                optional: false,
            }
        }
        Rule::empty_expression => {
            let arg = walk_expression(inner_nokw(pair).next().unwrap())?;
            ExprKind::Call {
                callee: Box::new(Expression::ident("empty")),
                args: vec![Argument::positional(arg)],
                optional: false,
            }
        }
        Rule::print_expression => {
            let arg = walk_expression(inner_nokw(pair).next().unwrap())?;
            ExprKind::Call {
                callee: Box::new(Expression::ident("__php_print_expr")),
                args: vec![Argument::positional(arg)],
                optional: false,
            }
        }
        Rule::include_operand => return walk_expression(pair.into_inner().next().unwrap()),
        Rule::include_expression => {
            let include_kind = php_include_kind(&pair)?;
            let arg = walk_expression(inner_nokw(pair).next().unwrap())?;
            build_php_dynamic_include_call(include_kind, arg, &span).kind
        }
        Rule::unset_expression => {
            // PHP `unset($a, $b, $obj->prop, $arr[$k])` — walker
            // rewrites each target individually based on its shape:
            //   - Ident:                    $x = null
            //   - Member ($obj->prop):     `__unset` magic dispatch,
            //                               else direct property delete
            //   - Index  ($arr[$k]):        `ecma:object.delete($arr, $k)`
            // Multi-arg unset becomes a Sequence of these. Walker
            // suppresses `__get` wrap on each arg via ASSIGN_LHS_DEPTH
            // so the LHS shape stays raw.
            let mut stmts: Vec<Expression> = Vec::new();
            for p in pair.into_inner() {
                if !matches!(p.as_rule(), Rule::expression) {
                    continue;
                }
                ASSIGN_LHS_DEPTH.with(|d| *d.borrow_mut() += 1);
                let walked = walk_expression(p);
                ASSIGN_LHS_DEPTH.with(|d| {
                    let mut bd = d.borrow_mut();
                    *bd = bd.saturating_sub(1);
                });
                let walked = walked?;
                stmts.push(build_unset_rewrite(walked, &span));
            }
            if stmts.is_empty() {
                ExprKind::Lit(Literal::Null)
            } else if stmts.len() == 1 {
                stmts.into_iter().next().unwrap().kind
            } else {
                ExprKind::Sequence(stmts)
            }
        }
        Rule::list_expression => {
            // list($a, $b, $c) / list('key' => $v) — destructure target.
            // Keep explicit holes for positional lists, and map keyed
            // entries onto the same object-pattern form used by `[...]`.
            let mut positional = Vec::new();
            let mut keyed = Vec::new();
            let mut saw_keyed = false;
            for p in pair.into_inner() {
                if matches!(p.as_rule(), Rule::list_element) {
                    let mut inner = p.into_inner();
                    let first = inner.next();
                    let second = inner.next();
                    match (first, second) {
                        (Some(key_pair), Some(value_pair)) => {
                            saw_keyed = true;
                            let key_expr = walk_expression(key_pair)?;
                            let value_expr = walk_expression(value_pair)?;
                            if let (Some(key), Some(value)) = (
                                literal_key_name(&key_expr),
                                expression_to_binding_pattern(&value_expr),
                            ) {
                                keyed.push(ObjectPatternProp {
                                    key,
                                    value: Some(value),
                                    default: None,
                                    is_rest: false,
                                });
                            }
                        }
                        (Some(e), None) => {
                            let expr = walk_expression(e)?;
                            if let Some(pattern) = expression_to_binding_pattern(&expr) {
                                positional.push(ArrayPatternElem::Pattern(pattern, None));
                            } else {
                                positional.push(ArrayPatternElem::Hole);
                            }
                        }
                        _ => positional.push(ArrayPatternElem::Hole),
                    }
                }
            }
            if saw_keyed {
                ExprKind::Destructure(DestructurePattern::Object(keyed))
            } else {
                ExprKind::Destructure(DestructurePattern::Array(positional))
            }
        }
        Rule::array_expression | Rule::short_array_expression => return walk_array(pair),
        Rule::closure_expression => return walk_closure(pair),
        Rule::arrow_function => return walk_arrow_function(pair),

        Rule::throw_expression => {
            // PHP 8: `throw` can appear in expression position (e.g.
            // `$x ?? throw new E`). Common AST has no Throw-expression
            // node; wrap as an immediately-invoked closure whose body
            // throws — produces the same control flow without leaking
            // PHP-specific shape into the compiler.
            let inner_expr = walk_expression(inner_nokw(pair).next().unwrap())?;
            let throw_stmt = Statement::with_span(
                StmtKind::Throw {
                    expr: Some(inner_expr),
                    cause: None,
                },
                span.clone(),
            );
            let lambda = Expression::with_span(
                ExprKind::Lambda {
                    params: vec![],
                    body: LambdaBody::Block(vec![throw_stmt]),
                    is_async: false,
                    captures: vec![],
                },
                span.clone(),
            );
            ExprKind::Call {
                callee: Box::new(lambda),
                args: vec![],
                optional: false,
            }
        }

        other => return Err(format!("walker: unhandled expression rule {:?}", other)),
    };

    Ok(Expression::with_span(kind, span))
}

fn php_include_kind(pair: &Pair<Rule>) -> Result<&'static str, String> {
    for inner in pair.clone().into_inner() {
        let kind = match inner.as_rule() {
            Rule::kw_include => Some("include"),
            Rule::kw_include_once => Some("include_once"),
            Rule::kw_require => Some("require"),
            Rule::kw_require_once => Some("require_once"),
            _ => None,
        };
        if let Some(kind) = kind {
            return Ok(kind);
        }
    }
    Err("missing PHP include kind".to_string())
}

fn build_php_dynamic_include_call(kind: &str, target: Expression, span: &Span) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::ident("__php_dynamic_include")),
            args: vec![
                Argument::positional(Expression::with_span(
                    ExprKind::Lit(Literal::Str(kind.to_string())),
                    span.clone(),
                )),
                Argument::positional(target),
            ],
            optional: false,
        },
        span.clone(),
    )
}

fn walk_left_assoc_binary(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner().peekable();
    let mut left = walk_expression(inner.next().unwrap())?;
    while let Some(op_pair) = inner.next() {
        let op_str = match op_pair.as_rule() {
            Rule::logic_or_op => "||".to_string(),
            Rule::logic_and_op => "&&".to_string(),
            Rule::bit_or_op => "|".to_string(),
            Rule::bit_xor_op => "^".to_string(),
            Rule::bit_and_op => "&".to_string(),
            _ => op_pair.as_str().to_string(),
        };
        let right_pair = match inner.next() {
            Some(p) => p,
            None => break,
        };
        let right = walk_expression(right_pair)?;
        let op = parse_binop(&op_str);
        let (l, r) = if op == BinOp::Concat {
            // PHP `$a . $b` invokes `__toString` on objects. Wrap each
            // operand in an IIFE that calls `__toString` when present:
            //   ((v) => v && v.__toString ? v.__toString() : v)($x)
            (
                php_concat_operand_coerce(left, &span),
                php_concat_operand_coerce(right, &span),
            )
        } else {
            (left, right)
        };
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(l),
                right: Box::new(r),
            },
            span.clone(),
        );
    }
    Ok(left)
}

fn php_concat_operand_coerce(expr: Expression, span: &Span) -> Expression {
    match &expr.kind {
        // The result of a prior concat is already a stringy value. Re-wrapping
        // the whole left subtree on every `.` step causes AST growth to explode
        // on long concat chains.
        ExprKind::Binary {
            op: BinOp::Concat, ..
        } => expr,
        _ => php_tostring_coerce(expr, span),
    }
}

/// Wrap `expr` in an IIFE that invokes `__toString()` when `expr` is an
/// object with that magic method, returns the value otherwise. PHP-only
/// — used for `.` concat and string coercion paths to satisfy the
/// `Stringable` contract without compiler-side runtime probes.
fn php_tostring_coerce(expr: Expression, span: &Span) -> Expression {
    let v_ident = || Expression::with_span(ExprKind::Ident("v".to_string()), span.clone());
    let ts_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(v_ident()),
            field: "__toString".to_string(),
            null_safe: false,
        },
        span.clone(),
    );
    // v && v.__toString
    let cond = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(v_ident()),
            right: Box::new(ts_member),
        },
        span.clone(),
    );
    // v.__toString()
    let call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(v_ident()),
                    field: "__toString".to_string(),
                    null_safe: false,
                },
                span.clone(),
            )),
            args: vec![],
            optional: false,
        },
        span.clone(),
    );
    // PHP semantics:
    //   true  → "1"
    //   false → ""
    //   null  → ""
    //   object with __toString → v.__toString()
    //   other → v (string/number pass through)
    let str1 = || Expression::with_span(ExprKind::Lit(Literal::Str("1".to_string())), span.clone());
    let str_empty =
        || Expression::with_span(ExprKind::Lit(Literal::Str(String::new())), span.clone());
    let bool_true = Expression::with_span(ExprKind::Lit(Literal::Bool(true)), span.clone());
    let null_lit = Expression::with_span(ExprKind::Lit(Literal::Null), span.clone());
    // if v === true → "1"
    let is_true = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(v_ident()),
            right: Box::new(bool_true),
        },
        span.clone(),
    );
    // if v === false || v === null → ""
    let bool_false = Expression::with_span(ExprKind::Lit(Literal::Bool(false)), span.clone());
    let is_false = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(v_ident()),
            right: Box::new(bool_false),
        },
        span.clone(),
    );
    let is_null = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(v_ident()),
            right: Box::new(null_lit),
        },
        span.clone(),
    );
    let is_falsy_lit = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(is_false),
            right: Box::new(is_null),
        },
        span.clone(),
    );
    // inner: if v && v.__toString → v.__toString() else v
    let inner = Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(call),
            else_: Box::new(v_ident()),
        },
        span.clone(),
    );
    // if is_falsy → "" else inner
    let level2 = Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(is_falsy_lit),
            then: Box::new(str_empty()),
            else_: Box::new(inner),
        },
        span.clone(),
    );
    // if is_true → "1" else level2
    let body = Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(is_true),
            then: Box::new(str1()),
            else_: Box::new(level2),
        },
        span.clone(),
    );
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Lambda {
                    params: vec![Param {
                        name: "v".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    }],
                    body: LambdaBody::Expr(Box::new(body)),
                    is_async: false,
                    captures: vec![],
                },
                span.clone(),
            )),
            args: vec![Argument::positional(expr)],
            optional: false,
        },
        span.clone(),
    )
}

fn parse_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "**" => BinOp::Pow,
        "." => BinOp::Concat,
        "==" => BinOp::Eq,
        "===" => BinOp::StrictEq,
        "!=" | "<>" => BinOp::NotEq,
        "!==" => BinOp::StrictNotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "<=>" => BinOp::Spaceship,
        "&&" | "and" | "AND" => BinOp::And,
        "||" | "or" | "OR" => BinOp::Or,
        "xor" | "XOR" => BinOp::Xor,
        "|" => BinOp::BitOr,
        "&" => BinOp::BitAnd,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "instanceof" | "INSTANCEOF" => BinOp::InstanceOf,
        "??" => BinOp::NullCoalesce,
        _ => BinOp::Add, // fallback — safer than panic
    }
}

/// Build the rewrite for a single `unset($target)` operation:
///   - `$x`         → `$x = null`
///   - `$obj->prop` → `($_t = $obj, typeof $_t->__unset === "function"
///                     ? $_t->__unset("prop") : ($_t->prop = null))`
///   - `$arr[$k]`   → `ecma:object.delete($arr, $k)`
fn build_unset_rewrite(target: Expression, span: &Span) -> Expression {
    match &target.kind {
        ExprKind::Ident(_) => Expression::with_span(
            ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(Expression::null()),
            },
            span.clone(),
        ),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } if !*null_safe && !field.starts_with("__") => {
            let obj = (**object).clone();
            let field = field.clone();
            let tmp = next_tmp_name("unset_recv");
            let tmp_ident = || Expression::with_span(ExprKind::Ident(tmp.clone()), span.clone());
            let save = Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(tmp_ident()),
                    value: Box::new(obj),
                },
                span.clone(),
            );
            let unset_member = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(tmp_ident()),
                    field: "__unset".to_string(),
                    null_safe: false,
                },
                span.clone(),
            );
            let has_unset = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        ExprKind::TypeOf(Box::new(unset_member)),
                        span.clone(),
                    )),
                    right: Box::new(Expression::string("function")),
                },
                span.clone(),
            );
            let magic_unset_call = Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(tmp_ident()),
                            field: "__unset".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    )),
                    args: vec![Argument::positional(Expression::string(&field))],
                    optional: false,
                },
                span.clone(),
            );
            let direct_delete = Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(tmp_ident()),
                            field: field.clone(),
                            null_safe: false,
                        },
                        span.clone(),
                    )),
                    value: Box::new(Expression::null()),
                },
                span.clone(),
            );
            let ternary = Expression::with_span(
                ExprKind::Ternary {
                    cond: Box::new(has_unset),
                    then: Box::new(magic_unset_call),
                    else_: Box::new(direct_delete),
                },
                span.clone(),
            );
            Expression::with_span(ExprKind::Sequence(vec![save, ternary]), span.clone())
        }
        ExprKind::Index { .. } => {
            // `unset($arr[$k])` → ExprKind::Delete (compiler routes
            // to `ecma:object.delete($arr, $k)`, polymorphic over
            // Array / Map / Ordinary backings).
            Expression::with_span(ExprKind::Delete(Box::new(target)), span.clone())
        }
        _ => {
            // Fallback: assign null. Best we can do without a real
            // l-value reference.
            Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(target),
                    value: Box::new(Expression::null()),
                },
                span.clone(),
            )
        }
    }
}

/// Build the magic-`__isset` rewrite for `isset($obj->prop)` checks:
///
///     ($_t = $obj,
///      typeof $_t->prop === "undefined" &&
///      typeof $_t->__isset === "function"
///        ? ($_t->__isset("prop") ? true : null)
///        : $_t->prop)
///
/// The inner `$_t->__isset(name) ? true : null` normalises the user's
/// `__isset` return value so the outer `isset(...)` host call (which
/// tests not-null-not-undefined) reports "set" / "not set" correctly.
fn build_magic_isset_rewrite(obj: Expression, field: String, span: &Span) -> Expression {
    let tmp = next_tmp_name("isset_recv");
    let tmp_ident = || Expression::with_span(ExprKind::Ident(tmp.clone()), span.clone());
    let save = Expression::with_span(
        ExprKind::Assign {
            target: Box::new(tmp_ident()),
            value: Box::new(obj),
        },
        span.clone(),
    );
    let direct_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(tmp_ident()),
            field: field.clone(),
            null_safe: false,
        },
        span.clone(),
    );
    let isset_member_chk = Expression::with_span(
        ExprKind::Member {
            object: Box::new(tmp_ident()),
            field: "__isset".to_string(),
            null_safe: false,
        },
        span.clone(),
    );
    let prop_missing = Expression::with_span(
        ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::In,
                    left: Box::new(Expression::string(&field)),
                    right: Box::new(tmp_ident()),
                },
                span.clone(),
            )),
        },
        span.clone(),
    );
    let has_isset = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::with_span(
                ExprKind::TypeOf(Box::new(isset_member_chk)),
                span.clone(),
            )),
            right: Box::new(Expression::string("function")),
        },
        span.clone(),
    );
    let cond = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(prop_missing),
            right: Box::new(has_isset),
        },
        span.clone(),
    );
    let magic_isset_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(tmp_ident()),
                    field: "__isset".to_string(),
                    null_safe: false,
                },
                span.clone(),
            )),
            args: vec![Argument::positional(Expression::string(&field))],
            optional: false,
        },
        span.clone(),
    );
    let normalized = Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(magic_isset_call),
            then: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
            else_: Box::new(Expression::null()),
        },
        span.clone(),
    );
    let ternary = Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(normalized),
            else_: Box::new(direct_member),
        },
        span.clone(),
    );
    Expression::with_span(ExprKind::Sequence(vec![save, ternary]), span.clone())
}

/// Build the magic-`__invoke` rewrite for `$var(args)` calls:
///
///     (typeof $var === "function"
///      || (typeof $var === "string" && function_exists($var)))
///       ? $var(args)
///       : $var->__invoke(args)
fn build_magic_invoke_rewrite(
    receiver: Expression,
    args: Vec<Argument>,
    span: &Span,
) -> Expression {
    let typeof_expr =
        Expression::with_span(ExprKind::TypeOf(Box::new(receiver.clone())), span.clone());
    let is_function = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(typeof_expr),
            right: Box::new(Expression::string("function")),
        },
        span.clone(),
    );
    let is_string = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::with_span(
                ExprKind::TypeOf(Box::new(receiver.clone())),
                span.clone(),
            )),
            right: Box::new(Expression::string("string")),
        },
        span.clone(),
    );
    let function_exists = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::ident("function_exists")),
            args: vec![Argument::positional(receiver.clone())],
            optional: false,
        },
        span.clone(),
    );
    let cond = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(is_function),
            right: Box::new(Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(is_string),
                    right: Box::new(function_exists),
                },
                span.clone(),
            )),
        },
        span.clone(),
    );
    let direct_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(receiver.clone()),
            args: args.clone(),
            optional: false,
        },
        span.clone(),
    );
    let invoke_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(receiver),
            field: "__invoke".to_string(),
            null_safe: false,
        },
        span.clone(),
    );
    let invoke_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(invoke_member),
            args,
            optional: false,
        },
        span.clone(),
    );
    Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(direct_call),
            else_: Box::new(invoke_call),
        },
        span.clone(),
    )
}

/// Build the magic-`__callStatic` rewrite for `Class::method(args)`:
///
///     typeof Class::method !== "function" &&
///     typeof Class.__callStatic === "function"
///       ? Class.__callStatic("method", [args])
///       : Class::method(args)
///
/// Uses Member-call shape (`Class.__callStatic(...)`) so the compiler's
/// static-method-call dispatch picks it up and prepends the class
/// object as `$this`.
fn build_magic_call_static_rewrite(
    class_expr: Expression,
    method_name: String,
    args: Vec<Argument>,
    span: &Span,
) -> Expression {
    // Use Member-shape (`Class.method(...)`) for both branches so the
    // compiler's static-method-on-user-class dispatch fires (calls.rs
    // ~600), which pushes the class object as `$this` slot 0 — load-
    // bearing for late-static-binding (`static::X` walked as
    // `$this::X` resolves the class const / static field on `$this`
    // when `$this` is the calling class).
    let direct_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(class_expr.clone()),
            field: method_name.clone(),
            null_safe: false,
        },
        span.clone(),
    );
    let direct_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(direct_member.clone()),
            args: args.clone(),
            optional: false,
        },
        span.clone(),
    );
    let static_call_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(class_expr.clone()),
            field: "__callStatic".to_string(),
            null_safe: false,
        },
        span.clone(),
    );
    let lacks_method = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::with_span(
                ExprKind::TypeOf(Box::new(direct_member)),
                span.clone(),
            )),
            right: Box::new(Expression::string("function")),
        },
        span.clone(),
    );
    let has_static = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::with_span(
                ExprKind::TypeOf(Box::new(static_call_member)),
                span.clone(),
            )),
            right: Box::new(Expression::string("function")),
        },
        span.clone(),
    );
    let cond = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(lacks_method),
            right: Box::new(has_static),
        },
        span.clone(),
    );
    let args_array = Expression::with_span(
        ExprKind::Array(
            args.iter()
                .map(|a| ArrayElement {
                    key: None,
                    value: a.value.clone(),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ),
        span.clone(),
    );
    let static_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(class_expr),
                    field: "__callStatic".to_string(),
                    null_safe: false,
                },
                span.clone(),
            )),
            args: vec![
                Argument::positional(Expression::string(&method_name)),
                Argument::positional(args_array),
            ],
            optional: false,
        },
        span.clone(),
    );
    Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(static_call),
            else_: Box::new(direct_call),
        },
        span.clone(),
    )
}

/// Build the magic-`__get` rewrite for `$obj->prop` reads:
///
///     ($_t = $obj,
///      typeof $_t->prop === "undefined" &&
///      typeof $_t->__get === "function"
///        ? $_t->__get("prop")
///        : $_t->prop)
fn build_magic_get_rewrite(obj: Expression, name: String, span: &Span) -> Expression {
    let tmp = next_tmp_name("get_recv");
    let tmp_ident = || Expression::with_span(ExprKind::Ident(tmp.clone()), span.clone());
    let save = Expression::with_span(
        ExprKind::Assign {
            target: Box::new(tmp_ident()),
            value: Box::new(obj),
        },
        span.clone(),
    );
    let direct_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(tmp_ident()),
            field: name.clone(),
            null_safe: false,
        },
        span.clone(),
    );
    let get_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(tmp_ident()),
            field: "__get".to_string(),
            null_safe: false,
        },
        span.clone(),
    );
    let prop_missing = Expression::with_span(
        ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::In,
                    left: Box::new(Expression::string(&name)),
                    right: Box::new(tmp_ident()),
                },
                span.clone(),
            )),
        },
        span.clone(),
    );
    let has_get = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::with_span(
                ExprKind::TypeOf(Box::new(get_member)),
                span.clone(),
            )),
            right: Box::new(Expression::string("function")),
        },
        span.clone(),
    );
    let cond = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(prop_missing),
            right: Box::new(has_get),
        },
        span.clone(),
    );
    let magic_get_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(tmp_ident()),
                    field: "__get".to_string(),
                    null_safe: false,
                },
                span.clone(),
            )),
            args: vec![Argument::positional(Expression::string(&name))],
            optional: false,
        },
        span.clone(),
    );
    let ternary = Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(magic_get_call),
            else_: Box::new(direct_member),
        },
        span.clone(),
    );
    Expression::with_span(ExprKind::Sequence(vec![save, ternary]), span.clone())
}

/// Build the magic-`__call` rewrite for `$obj->method(args)` invocation:
///
///     ($_t = $obj,
///      typeof $_t->method !== "function" &&
///      typeof $_t->__call === "function"
///        ? $_t->__call("method", [args])
///        : $_t->method(args))
///
/// Extracted out of `apply_postfix` so that walker's recursion frame
/// stays small — the rewrite allocates ~20 Expression nodes, and
/// nested-closure / chained-call tests would blow the test-thread
/// stack with all those locals live in one frame.
fn build_magic_call_rewrite(
    member_object: Expression,
    method_name: String,
    args: Vec<Argument>,
    span: &Span,
) -> Expression {
    let tmp = next_tmp_name("call_recv");
    let tmp_ident = || Expression::with_span(ExprKind::Ident(tmp.clone()), span.clone());
    let save = Expression::with_span(
        ExprKind::Assign {
            target: Box::new(tmp_ident()),
            value: Box::new(member_object),
        },
        span.clone(),
    );
    let direct_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(tmp_ident()),
            field: method_name.clone(),
            null_safe: false,
        },
        span.clone(),
    );
    let regular_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(direct_member.clone()),
            args: args.clone(),
            optional: false,
        },
        span.clone(),
    );
    let call_member_for_check = Expression::with_span(
        ExprKind::Member {
            object: Box::new(tmp_ident()),
            field: "__call".to_string(),
            null_safe: false,
        },
        span.clone(),
    );
    let has_call = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::with_span(
                ExprKind::TypeOf(Box::new(call_member_for_check)),
                span.clone(),
            )),
            right: Box::new(Expression::string("function")),
        },
        span.clone(),
    );
    let lacks_method = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::with_span(
                ExprKind::TypeOf(Box::new(direct_member.clone())),
                span.clone(),
            )),
            right: Box::new(Expression::string("function")),
        },
        span.clone(),
    );
    let cond = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(lacks_method),
            right: Box::new(has_call),
        },
        span.clone(),
    );
    let args_array = Expression::with_span(
        ExprKind::Array(
            args.iter()
                .map(|a| ArrayElement {
                    key: None,
                    value: a.value.clone(),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ),
        span.clone(),
    );
    let call_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(tmp_ident()),
            field: "__call".to_string(),
            null_safe: false,
        },
        span.clone(),
    );
    let magic_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(call_member),
            args: vec![
                Argument::positional(Expression::string(&method_name)),
                Argument::positional(args_array),
            ],
            optional: false,
        },
        span.clone(),
    );
    let ternary = Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(magic_call),
            else_: Box::new(regular_call),
        },
        span.clone(),
    );
    Expression::with_span(ExprKind::Sequence(vec![save, ternary]), span.clone())
}

/// Build the magic-`__set` rewrite for a `$obj->prop = $val` assignment:
///
///     ($_t = $obj, $_v = $val,
///      (typeof $_t->prop === "undefined" &&
///       typeof $_t->__set === "function")
///        ? $_t->__set("prop", $_v)
///        : ($_t->prop = $_v))
///
/// Extracted out of `walk_assignment` so the latter's stack frame stays
/// small — the rewrite allocates ~15 Expression nodes, and currying /
/// nested-closure tests would blow the test-thread stack if all those
/// locals were live at every recursive walk_expression frame.
fn build_magic_set_rewrite(
    obj: Expression,
    field: String,
    rhs: Expression,
    span: &Span,
) -> Expression {
    let tmp_recv = next_tmp_name("set_recv");
    let tmp_val = next_tmp_name("set_val");
    let recv_ident = || Expression::with_span(ExprKind::Ident(tmp_recv.clone()), span.clone());
    let val_ident = || Expression::with_span(ExprKind::Ident(tmp_val.clone()), span.clone());
    let save_recv = Expression::with_span(
        ExprKind::Assign {
            target: Box::new(recv_ident()),
            value: Box::new(obj),
        },
        span.clone(),
    );
    let save_val = Expression::with_span(
        ExprKind::Assign {
            target: Box::new(val_ident()),
            value: Box::new(rhs),
        },
        span.clone(),
    );
    let direct_member = Expression::with_span(
        ExprKind::Member {
            object: Box::new(recv_ident()),
            field: field.clone(),
            null_safe: false,
        },
        span.clone(),
    );
    let set_member_for_check = Expression::with_span(
        ExprKind::Member {
            object: Box::new(recv_ident()),
            field: "__set".to_string(),
            null_safe: false,
        },
        span.clone(),
    );
    let prop_missing = Expression::with_span(
        ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::In,
                    left: Box::new(Expression::string(&field)),
                    right: Box::new(recv_ident()),
                },
                span.clone(),
            )),
        },
        span.clone(),
    );
    let has_set = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::with_span(
                ExprKind::TypeOf(Box::new(set_member_for_check)),
                span.clone(),
            )),
            right: Box::new(Expression::string("function")),
        },
        span.clone(),
    );
    let cond = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(prop_missing),
            right: Box::new(has_set),
        },
        span.clone(),
    );
    let magic_set_call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(recv_ident()),
                    field: "__set".to_string(),
                    null_safe: false,
                },
                span.clone(),
            )),
            args: vec![
                Argument::positional(Expression::string(&field)),
                Argument::positional(val_ident()),
            ],
            optional: false,
        },
        span.clone(),
    );
    let direct_assign = Expression::with_span(
        ExprKind::Assign {
            target: Box::new(direct_member),
            value: Box::new(val_ident()),
        },
        span.clone(),
    );
    let ternary = Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(magic_set_call),
            else_: Box::new(direct_assign),
        },
        span.clone(),
    );
    Expression::with_span(
        ExprKind::Sequence(vec![save_recv, save_val, ternary]),
        span.clone(),
    )
}

fn walk_assignment(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let lhs_pair = inner.next().unwrap();
    // If lhs_pair is yield_expression, just walk it through.
    if matches!(lhs_pair.as_rule(), Rule::yield_expression) {
        return walk_expression(lhs_pair);
    }
    // Mark that we're walking the LHS of an `=` so the property-access
    // walker can suppress the magic-`__get` ternary on the OUTERMOST
    // chain op (that op is the WRITE target, not a read). Inner reads
    // in a chain like `$a->b->c = $val` (where `->b` is a read) still
    // get the magic dispatch.
    //
    // Only set the flag when an `=` operator actually follows the LHS
    // pair — pest's `assignment_expression` grammar wraps every
    // expression so a bare `echo $obj->prop;` would otherwise look
    // like an assignment LHS.
    let has_assign_op = inner.peek().is_some();
    if has_assign_op {
        ASSIGN_LHS_DEPTH.with(|d| *d.borrow_mut() += 1);
    }
    let lhs_result = walk_expression(lhs_pair);
    if has_assign_op {
        ASSIGN_LHS_DEPTH.with(|d| {
            let mut bd = d.borrow_mut();
            *bd = bd.saturating_sub(1);
        });
    }
    let lhs_walked = lhs_result?;
    // Defer destructure conversion until we confirm there's actually a `=`.
    // The grammar wraps every expression in `assignment_expression`, so
    // an isolated `[]` (without `=`) reaches walk_assignment too — if we
    // converted eagerly, RHS empty arrays would compile to `Op::NULL`
    // via the `ExprKind::Destructure` arm.
    let has_assign = inner.peek().is_some();
    let lhs = if has_assign {
        expression_into_destructure_target(lhs_walked)
    } else {
        lhs_walked
    };
    if let Some(op_pair) = inner.next() {
        let op = op_pair.as_str();
        let rhs = walk_expression(inner.next().unwrap())?;
        // PHP `__set` magic method: when `$obj->prop = $val` is
        // executed and `prop` isn't an own property of `$obj`, PHP
        // dispatches to `$obj->__set("prop", $val)` if the class
        // defines it. Walker rewrites simple-target Assigns:
        //
        // Note: walker is conservative about WHEN to wrap to avoid
        // explosive AST depth in chained / deeply nested forms.
        //
        //   $obj->prop = $val
        //     →
        //   ($_t = $obj, $_v = $val,
        //    (typeof $_t->prop === "undefined" &&
        //     typeof $_t->__set === "function")
        //       ? $_t->__set("prop", $_v)
        //       : ($_t->prop = $_v))
        //
        // Skipped for `__`-prefixed names (internals like `__type`),
        // null-safe member access, compound-op assignments (those go
        // through Read+Op+Write semantics where the read also routes
        // via `__get` already), and non-Member targets.
        if op == "=" {
            if let ExprKind::Member {
                object,
                field,
                null_safe,
            } = &lhs.kind
            {
                if !null_safe && !field.starts_with("__") && !is_php_this_expr(object) {
                    let obj = (**object).clone();
                    let field = field.clone();
                    return Ok(build_magic_set_rewrite(obj, field, rhs.clone(), &span));
                }
            }
        }
        let kind = match op {
            "=" => ExprKind::Assign {
                target: Box::new(lhs),
                value: Box::new(rhs),
            },
            other => {
                let cop = parse_compound_op(other);
                // CompoundAssign is a stmt-level node; expression-level
                // compound assignments synthesize `target = target OP rhs`.
                let combined = Expression::with_span(
                    ExprKind::Binary {
                        op: compound_to_binop(cop),
                        left: Box::new(lhs.clone()),
                        right: Box::new(rhs),
                    },
                    span.clone(),
                );
                ExprKind::Assign {
                    target: Box::new(lhs),
                    value: Box::new(combined),
                }
            }
        };
        Ok(Expression::with_span(kind, span))
    } else {
        Ok(lhs)
    }
}

fn parse_compound_op(s: &str) -> CompoundOp {
    match s {
        "+=" => CompoundOp::Add,
        "-=" => CompoundOp::Sub,
        "*=" => CompoundOp::Mul,
        "/=" => CompoundOp::Div,
        "%=" => CompoundOp::Mod,
        ".=" => CompoundOp::Concat,
        "**=" => CompoundOp::Pow,
        "<<=" => CompoundOp::Shl,
        ">>=" => CompoundOp::Shr,
        "&=" => CompoundOp::BitAnd,
        "|=" => CompoundOp::BitOr,
        "^=" => CompoundOp::BitXor,
        "&&=" => CompoundOp::And,
        "||=" => CompoundOp::Or,
        "??=" => CompoundOp::NullCoalesce,
        _ => CompoundOp::Add,
    }
}

fn compound_to_binop(op: CompoundOp) -> BinOp {
    match op {
        CompoundOp::Add => BinOp::Add,
        CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul,
        CompoundOp::Div => BinOp::Div,
        CompoundOp::Mod => BinOp::Mod,
        CompoundOp::Pow => BinOp::Pow,
        CompoundOp::Concat => BinOp::Concat,
        CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr,
        CompoundOp::UShr => BinOp::UShr,
        CompoundOp::BitAnd => BinOp::BitAnd,
        CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor,
        CompoundOp::And => BinOp::And,
        CompoundOp::Or => BinOp::Or,
        CompoundOp::NullCoalesce => BinOp::NullCoalesce,
        CompoundOp::IDiv => BinOp::IDiv,
    }
}

fn walk_yield(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    // Detect `yield from` from the source slice — kw_yield/kw_yield_from
    // are filtered out alongside the rest of the keyword tokens.
    let yield_from = pair
        .as_str()
        .trim_start()
        .to_lowercase()
        .starts_with("yield from");
    let mut inner = inner_nokw(pair);
    if yield_from {
        let val = walk_expression(inner.next().unwrap())?;
        return Ok(Expression::with_span(
            ExprKind::YieldFrom(Box::new(val)),
            span,
        ));
    }
    // bare `yield`, `yield expr`, or `yield key => value`
    let remaining: Vec<Pair<Rule>> = inner.collect();
    let val = match remaining.as_slice() {
        [] => None,
        [value] => Some(walk_expression(value.clone())?),
        [key, value] => Some(Expression::with_span(
            ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("__vybe_generator_yield"),
                    value: Expression::bool(true),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("key"),
                    value: walk_expression(key.clone())?,
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("value"),
                    value: walk_expression(value.clone())?,
                },
            ]),
            span,
        )),
        _ => return Err("unsupported yield expression shape".to_string()),
    };
    Ok(Expression::with_span(
        ExprKind::Yield(val.map(Box::new)),
        span,
    ))
}

fn walk_ternary(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().unwrap())?;
    let next = inner.next();
    if next.is_none() {
        return Ok(cond);
    }
    let mut next = next.unwrap();
    // Two forms:
    //   `cond ? then : else`
    //   `cond ?: else` (Elvis — short ternary)
    if matches!(next.as_rule(), Rule::expression) {
        // We have a `then` branch.
        let then_expr = walk_expression(next)?;
        let else_expr = walk_expression(inner.next().unwrap())?;
        return Ok(Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(then_expr),
                else_: Box::new(else_expr),
            },
            span,
        ));
    } else {
        // Elvis: cond ?: else  →  cond ?? else (semantically close enough)
        // Actually PHP's ?: returns cond if truthy, else the right side.
        // We model as: cond ? cond : else
        next = inner.next().unwrap_or(next);
        let else_expr = walk_expression(next)?;
        Ok(Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(cond.clone()),
                then: Box::new(cond),
                else_: Box::new(else_expr),
            },
            span,
        ))
    }
}

fn walk_unary(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    if matches!(first.as_rule(), Rule::unary_op) {
        let op_str = first.as_str();
        // Normalise PHP `++` / `--` to a language-neutral call + assign
        // so the compiler never sees PHP-specific increment semantics.
        // Mirrors the postfix rewrite in `apply_postfix` — see that
        // comment for the full rationale.
        if op_str == "++" || op_str == "--" {
            let expr = walk_expression_as_assign_target(inner.next().unwrap())?;
            let helper = if op_str == "++" {
                "__php_increment"
            } else {
                "__php_decrement"
            };
            let callee = Expression::with_span(ExprKind::Ident(helper.to_string()), span.clone());
            let call = Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: vec![Argument::positional(expr.clone())],
                    optional: false,
                },
                span.clone(),
            );
            return Ok(Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(expr),
                    value: Box::new(call),
                },
                span,
            ));
        }
        let expr = walk_expression(inner.next().unwrap())?;
        let op = parse_unary_op(op_str);
        Ok(Expression::with_span(
            ExprKind::Unary {
                op,
                expr: Box::new(expr),
            },
            span,
        ))
    } else {
        walk_expression(first)
    }
}

fn parse_unary_op(s: &str) -> UnaryOp {
    match s {
        "!" => UnaryOp::Not,
        "~" => UnaryOp::BitNot,
        "-" => UnaryOp::Neg,
        "+" => UnaryOp::Pos,
        "++" => UnaryOp::PreInc,
        "--" => UnaryOp::PreDec,
        "@" => UnaryOp::Pos, // PHP error suppression — semantically a no-op for us
        "&" => UnaryOp::AddrOf,
        _ => UnaryOp::Pos,
    }
}

fn walk_php_variable_expr(pair: Pair<Rule>, span: Span) -> Result<Expression, String> {
    let raw = pair.as_str();
    if let Some(rest) = raw.strip_prefix("$$") {
        let key_expr = if rest.starts_with('$') {
            Expression::with_span(ExprKind::Ident(rest.to_string()), span.clone())
        } else {
            Expression::with_span(ExprKind::Lit(Literal::Str(rest.to_string())), span.clone())
        };
        return Ok(Expression::with_span(
            ExprKind::Index {
                object: Box::new(Expression::with_span(
                    ExprKind::Ident("__php_var_vars".to_string()),
                    span.clone(),
                )),
                index: Box::new(key_expr),
                null_safe: false,
            },
            span,
        ));
    }

    let mut inner = pair.clone().into_inner();
    if let Some(first) = inner.next() {
        if matches!(first.as_rule(), Rule::expression) {
            let key_expr = walk_expression(first)?;
            return Ok(Expression::with_span(
                ExprKind::Index {
                    object: Box::new(Expression::with_span(
                        ExprKind::Ident("__php_var_vars".to_string()),
                        span.clone(),
                    )),
                    index: Box::new(key_expr),
                    null_safe: false,
                },
                span,
            ));
        }
    }

    Ok(Expression::with_span(
        ExprKind::Ident(strip_dollar(raw).to_string()),
        span,
    ))
}

fn walk_property_hooks(
    pair: Pair<Rule>,
    type_hint: &Option<String>,
) -> Result<(Option<Vec<Statement>>, Option<PropertySetter>), String> {
    let mut getter = None;
    let mut setter = None;

    for hook in pair.into_inner() {
        match hook.as_rule() {
            Rule::property_get_hook => {
                if let Some(block) = hook
                    .into_inner()
                    .find(|p| matches!(p.as_rule(), Rule::block_statement))
                {
                    getter = Some(walk_statement_into_body(block)?);
                }
            }
            Rule::property_set_hook => {
                let mut param = Param {
                    name: "value".to_string(),
                    type_hint: type_hint.clone(),
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                };
                let mut body = Vec::new();
                for item in hook.into_inner() {
                    match item.as_rule() {
                        Rule::param_list => {
                            if let Some(first) = walk_params(item)?.into_iter().next() {
                                param = first;
                            }
                        }
                        Rule::block_statement => {
                            body = walk_statement_into_body(item)?;
                        }
                        _ => {}
                    }
                }
                setter = Some(PropertySetter { param, body });
            }
            _ => {}
        }
    }

    Ok((getter, setter))
}

fn walk_cast(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cast_kw = inner.next().unwrap().as_str().to_string();
    let expr = walk_expression(inner.next().unwrap())?;
    // PHP casts have defined runtime semantics (truncate toward zero for
    // `(int)`, ECMA `Boolean()`-shaped truthiness for `(bool)`, etc.) —
    // the vybex `ExprKind::Cast` case is compiled as a no-op. Normalise
    // at walker time into the PHP builtin call that already does the
    // right thing, so the compiler never needs PHP-cast awareness.
    let helper = match cast_kw
        .to_lowercase()
        .trim_start_matches('(')
        .trim_end_matches(')')
        .trim()
    {
        "int" | "integer" | "long" => Some("intval"),
        "float" | "double" | "real" => Some("floatval"),
        "bool" | "boolean" => Some("boolval"),
        "string" | "binary" => Some("strval"),
        // `(array)`, `(object)`, `(unset)` fall through to a Cast node;
        // the compiler currently keeps those as identity — if one of
        // those ever needs real semantics, handle it here too.
        _ => None,
    };
    if let Some(name) = helper {
        // PHP `(string) $x` invokes `__toString` if `$x` is an object
        // implementing Stringable. Wrap the operand for the string
        // cast so the host fn receives the coerced value.
        let arg_expr = if name == "strval" {
            php_tostring_coerce(expr, &span)
        } else {
            expr
        };
        let callee = Expression::with_span(ExprKind::Ident(name.to_string()), span.clone());
        return Ok(Expression::with_span(
            ExprKind::Call {
                callee: Box::new(callee),
                args: vec![Argument::positional(arg_expr)],
                optional: false,
            },
            span,
        ));
    }
    Ok(Expression::with_span(
        ExprKind::Cast {
            expr: Box::new(expr),
            type_name: cast_kw,
        },
        span,
    ))
}

fn walk_postfix(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let primary = inner.next().unwrap();
    // Track whether the very first chain element was a `$variable`
    // primary. This is used by `apply_postfix` to detect the
    // `$obj(args)` shape — invoking a value held in a variable —
    // which in PHP must dispatch through `__invoke` if the value
    // happens to be an object with that magic method. Function names
    // (bare identifiers) flow through the regular Call path; variables
    // need the wrapper.
    //
    // The primary pair is `Rule::primary_expression` (a non-silent
    // wrapper around alternatives like `variable | qualified_name |
    // …`). Peek at its first inner child to find the actual
    // primary kind.
    let is_var_primary = if matches!(primary.as_rule(), Rule::primary_expression) {
        primary
            .clone()
            .into_inner()
            .next()
            .map(|p| matches!(p.as_rule(), Rule::variable))
            .unwrap_or(false)
    } else {
        matches!(primary.as_rule(), Rule::variable)
    };
    // Suppress the magic-`__get` ternary wrap on the LAST op when
    // we're at the outermost level walking an Assign LHS. Walker
    // clears the flag here so nested expressions inside the postfix
    // chain don't see the suppression — only the LAST op being
    // applied (the assignment target) skips the wrap.
    //
    // The flag is read non-mutably to avoid leaving a 0 depth that
    // would let nested expressions accidentally claim themselves as
    // LHS targets. `is_assign_target` only fires when this walk_postfix
    // call is the OUTERMOST on the LHS (depth>0 AND we're at the
    // primary's chain) — nested walk_expression calls inside arg
    // lists / index expressions reset the depth below.
    let lhs_depth_was = ASSIGN_LHS_DEPTH.with(|d| *d.borrow());
    // Clear the depth while walking the primary + ops so nested
    // expressions inside (e.g. method args, indexes) don't inherit it.
    if lhs_depth_was > 0 {
        ASSIGN_LHS_DEPTH.with(|d| *d.borrow_mut() = 0);
    }
    let mut expr = walk_expression(primary)?;
    let mut from_variable = is_var_primary;
    let ops: Vec<_> = inner.collect();
    let n = ops.len();
    for (i, op_pair) in ops.iter().cloned().enumerate() {
        let is_last_op = i == n - 1;
        let next_is_inc_dec = ops
            .get(i + 1)
            .and_then(postfix_rule_kind)
            .is_some_and(|rule| matches!(rule, Rule::inc_dec_op));
        let is_assign_target = (lhs_depth_was > 0 && is_last_op) || next_is_inc_dec;
        expr = apply_postfix(expr, op_pair, &span, from_variable, is_assign_target)?;
        // After the first postfix is applied, the chain is a
        // computed value (member access, call result, etc.), not the
        // original variable any more.
        from_variable = false;
    }
    // Restore the depth so the outer walk_assignment sees the same
    // value it originally set.
    if lhs_depth_was > 0 {
        ASSIGN_LHS_DEPTH.with(|d| *d.borrow_mut() = lhs_depth_was);
    }
    Ok(expr)
}

fn apply_postfix(
    receiver: Expression,
    op: Pair<Rule>,
    span: &Span,
    from_variable: bool,
    is_assign_target: bool,
) -> Result<Expression, String> {
    // The grammar wraps all variants in a non-silent `postfix_op` rule, so
    // pest yields a `postfix_op` pair whose single child is the actual
    // op rule (`method_call_op`, `inc_dec_op`, etc.). Unwrap once so the
    // match below sees the real rule; otherwise every postfix silently
    // falls through to the `_ => Ok(receiver)` arm (dropping `$i++`,
    // `$obj->foo(...)`, `$arr[0]`, …).
    let op = if matches!(op.as_rule(), Rule::postfix_op) {
        op.into_inner().next().ok_or("empty postfix_op")?
    } else {
        op
    };
    let rule = op.as_rule();
    match rule {
        Rule::method_call_op => {
            // The grammar emits: ("?->"|"->") ~ member_name ~ "(" ~ arg_list? ~ ")"
            // The literal "->" / "?->" appears as a non-rule token, so
            // pest does NOT yield it as a child pair. Detect null-safe
            // from the outer pair's source text instead of trying to
            // read it from inner pairs.
            let null_safe = op.as_str().trim_start().starts_with("?->");
            let mut name_pair: Option<Pair<Rule>> = None;
            let mut arg_list_pair: Option<Pair<Rule>> = None;
            let mut is_fcc = false;
            for p in op.into_inner() {
                match p.as_rule() {
                    Rule::member_name => name_pair = Some(p),
                    Rule::arg_list => arg_list_pair = Some(p),
                    Rule::first_class_callable_op => {
                        is_fcc = true;
                    }
                    _ => {}
                }
            }
            let name_inner = name_pair
                .ok_or("method_call_op: missing name")?
                .into_inner()
                .next()
                .unwrap();
            if matches!(name_inner.as_rule(), Rule::variable | Rule::expression) {
                let member = Expression::with_span(
                    ExprKind::Index {
                        object: Box::new(receiver),
                        index: Box::new(walk_expression(name_inner)?),
                        null_safe,
                    },
                    span.clone(),
                );
                if is_fcc {
                    return Ok(php_first_class_callable_lambda(member, null_safe, span));
                }
                let args = arg_list_pair
                    .map(walk_args)
                    .transpose()?
                    .unwrap_or_default();
                return Ok(Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(member),
                        args,
                        optional: null_safe,
                    },
                    span.clone(),
                ));
            }
            let name = name_inner.as_str().to_string();
            // PHP exception accessor methods → property reads. Vybe's
            // exception ctor stamps `message`, `code`, etc. as plain
            // fields; PHP idiom is `$e->getMessage()`. Rewrite the
            // common method names to direct property access so the
            // existing field shape works without an Exception base
            // class with these methods defined.
            if !is_fcc && arg_list_pair.is_none() {
                // (field, default-when-absent). PHP getters with a defined
                // default get `field ?? default` so an absent field (Undefined)
                // becomes the PHP value: getPrevious()→null (so
                // `while ($e->getPrevious() !== null)` terminates) and
                // getCode()→0. Works cross-language too — a JS/Python exception
                // with no cause/code still yields the PHP default in PHP.
                let prop: Option<(&str, Option<Expression>)> = match name.as_str() {
                    "getMessage" => Some(("message", None)),
                    "getCode" => Some(("code", Some(Expression::int(0)))),
                    "getFile" => Some(("file", None)),
                    "getLine" => Some(("line", None)),
                    "getTrace" => Some(("trace", None)),
                    "getTraceAsString" => Some(("stack", None)),
                    "getPrevious" => Some(("cause", Some(Expression::null()))),
                    _ => None,
                };
                if let Some((field, default)) = prop {
                    let member = Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(receiver),
                            field: field.to_string(),
                            null_safe,
                        },
                        span.clone(),
                    );
                    let expr = match default {
                        Some(d) => Expression::with_span(
                            ExprKind::NullCoalesce {
                                left: Box::new(member),
                                right: Box::new(d),
                            },
                            span.clone(),
                        ),
                        None => member,
                    };
                    return Ok(expr);
                }
            }
            // PHP `Fiber` instance methods → bytecode adapter calls
            // that emit the VM's stack-switching ops (`RESUME`,
            // continuation property reads). Same shape as the
            // DateTime adapter — `$fiber->X(args)` rewrites to
            // `__php_fiber_X($fiber, args)`. The walker can't tell
            // a real Fiber from a user class with these method names,
            // so this rewrite intercepts unconditionally; users who
            // need their own `start`/`resume` should rename.
            if !is_fcc {
                // Only Fiber-specific names that don't collide with
                // PHP generators' API (`current` / `next` / `send` /
                // `getReturn` / `valid` are Generator methods, not
                // Fiber-specific). `start` / `resume` exist only on
                // Fibers; the four `isXxx` predicates are also
                // Fiber-only. Keeping `getReturn` out of this list
                // means user code calling `$generator->getReturn()`
                // routes through the VM's native generator dispatch
                // (which actually returns the generator's return
                // value); Fiber's `getReturn` is a TODO until VM
                // exposes a way to read continuation return.
                let fiber_target: Option<&str> = match name.as_str() {
                    "start" => Some("__php_fiber_start"),
                    "resume" => Some("__php_fiber_resume"),
                    "isStarted" => Some("__php_fiber_is_started"),
                    "isSuspended" => Some("__php_fiber_is_suspended"),
                    "isRunning" => Some("__php_fiber_is_running"),
                    "isTerminated" => Some("__php_fiber_is_terminated"),
                    _ => None,
                };
                if let Some(fname) = fiber_target {
                    let mut call_args: Vec<Argument> = vec![Argument::positional(receiver.clone())];
                    if let Some(al) = arg_list_pair.clone() {
                        call_args.extend(walk_args(al)?);
                    }
                    return Ok(Expression::with_span(
                        ExprKind::Call {
                            callee: Box::new(Expression::with_span(
                                ExprKind::Ident(fname.to_string()),
                                span.clone(),
                            )),
                            args: call_args,
                            optional: false,
                        },
                        span.clone(),
                    ));
                }
            }
            // PHP DateTime / DateTimeImmutable instance methods →
            // bytecode adapter calls (see emitter/php/datetime_adapter.rs).
            // Rewrites `$dt->X(...)` to `__php_dt_X($dt, ...)` which the
            // PHP profile binds to the corresponding `common:php.X`
            // emit target. Note: this runs unconditionally — user
            // classes that define `format`/`modify`/`diff`/etc. would
            // be rerouted; the trade-off is the same one the exception
            // accessor rewrite above accepts.
            if !is_fcc {
                let target_fn: Option<&str> = match name.as_str() {
                    "format" => Some("__php_dt_format"),
                    "getTimestamp" => Some("__php_dt_get_timestamp"),
                    "modify" => Some("__php_dt_modify"),
                    "diff" => Some("__php_dt_diff"),
                    "add" => Some("__php_dt_add"),
                    "sub" => Some("__php_dt_sub"),
                    _ => None,
                };
                if let Some(fname) = target_fn {
                    let mut call_args: Vec<Argument> = vec![Argument::positional(receiver.clone())];
                    if let Some(al) = arg_list_pair.clone() {
                        call_args.extend(walk_args(al)?);
                    }
                    // `$dt->format("LITERAL")` — pre-parse the format
                    // string at compile time and emit AST that calls
                    // ECMA Date getters (`getFullYear` / `getMonth` /
                    // `getDate` / etc.) with `padStart` for zero-
                    // padding. Bypasses any non-ECMA host fn.
                    if fname == "__php_dt_format" && call_args.len() == 2 {
                        if let ExprKind::Lit(Literal::Str(fmt)) = &call_args[1].value.kind {
                            let dt_expr = call_args[0].value.clone();
                            if let Some(formatted) =
                                crate::emitter::php::datetime_adapter::format_php_literal_to_ast(
                                    fmt, &dt_expr, &span,
                                )
                            {
                                return Ok(formatted);
                            }
                        }
                    }
                    // `$dt->modify("LITERAL")` — pre-parse the relative
                    // delta at compile time and route to the unit-
                    // specific adapter (`__php_dt_add_months`,
                    // `__php_dt_add_days`, ...) so the bytecode uses
                    // ECMA-spec `set<Component>` setters for calendar
                    // shifts and pure ms arithmetic for fixed-duration
                    // shifts. No runtime relative-string parser
                    // required.
                    if fname == "__php_dt_modify" && call_args.len() == 2 {
                        if let ExprKind::Lit(Literal::Str(s)) = &call_args[1].value.kind {
                            if let Some((n, unit)) =
                                crate::emitter::php::datetime_adapter::parse_relative_delta(s)
                            {
                                let adapter = match unit {
                                    "second" => "__php_dt_add_seconds",
                                    "minute" => "__php_dt_add_minutes",
                                    "hour" => "__php_dt_add_hours",
                                    "day" => "__php_dt_add_days",
                                    "week" => "__php_dt_add_weeks",
                                    "month" => "__php_dt_add_months",
                                    "year" => "__php_dt_add_years",
                                    _ => unreachable!(),
                                };
                                return Ok(Expression::with_span(
                                    ExprKind::Call {
                                        callee: Box::new(Expression::with_span(
                                            ExprKind::Ident(adapter.to_string()),
                                            span.clone(),
                                        )),
                                        args: vec![
                                            call_args.remove(0),
                                            Argument::positional(Expression::int(n)),
                                        ],
                                        optional: false,
                                    },
                                    span.clone(),
                                ));
                            }
                        }
                    }
                    return Ok(Expression::with_span(
                        ExprKind::Call {
                            callee: Box::new(Expression::with_span(
                                ExprKind::Ident(fname.to_string()),
                                span.clone(),
                            )),
                            args: call_args,
                            optional: false,
                        },
                        span.clone(),
                    ));
                }
            }
            let member = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(receiver),
                    field: name,
                    null_safe,
                },
                span.clone(),
            );
            if is_fcc {
                return Ok(php_first_class_callable_lambda(member, null_safe, span));
            }
            let args = arg_list_pair
                .map(walk_args)
                .transpose()?
                .unwrap_or_default();
            //
            //   $obj->method(a, b)
            //     →
            //   ($_t = $obj,
            //    typeof $_t->method === "function"
            //      ? $_t->method(a, b)
            //      : $_t->__call("method", [a, b]))
            //
            // The temp variable caches the receiver to avoid re-
            // evaluating side effects (computed property access etc.).
            // Skipped for null-safe (`?->`) — those keep the original
            // null-short-circuit semantics — and for the
            // already-rewritten exception accessor / DateTime adapter
            // forms (those return early above this branch).
            if null_safe {
                return Ok(Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(member),
                        args,
                        optional: null_safe,
                    },
                    span.clone(),
                ));
            }
            // `member` was built above as `Member { object: receiver, field: name, .. }`.
            // Extract field name from `member` for use in __call's literal arg.
            let (member_object, method_name) = match member.kind.clone() {
                ExprKind::Member { object, field, .. } => (*object, field),
                other => {
                    return Ok(Expression::with_span(
                        ExprKind::Call {
                            callee: Box::new(Expression::with_span(other, member.span.clone())),
                            args,
                            optional: null_safe,
                        },
                        span.clone(),
                    ));
                }
            };
            let directory_adapter: Option<&str> = match method_name.as_str() {
                "read" => Some("__php_dir_read"),
                "close" => Some("__php_dir_close"),
                _ => None,
            };
            if let Some(adapter_name) = directory_adapter {
                let mut adapter_args = vec![Argument::positional(member_object.clone())];
                adapter_args.extend(args.clone());
                let directory_type = Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(member_object.clone()),
                        field: "__type".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                );
                let is_directory = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(directory_type),
                        right: Box::new(Expression::string("Directory")),
                    },
                    span.clone(),
                );
                let adapter_call = Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(Expression::with_span(
                            ExprKind::Ident(adapter_name.to_string()),
                            span.clone(),
                        )),
                        args: adapter_args,
                        optional: false,
                    },
                    span.clone(),
                );
                let direct_member = Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(member_object.clone()),
                        field: method_name.clone(),
                        null_safe,
                    },
                    span.clone(),
                );
                let direct_call = Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(direct_member),
                        args: args.clone(),
                        optional: false,
                    },
                    span.clone(),
                );
                return Ok(Expression::with_span(
                    ExprKind::Ternary {
                        cond: Box::new(is_directory),
                        then: Box::new(adapter_call),
                        else_: Box::new(direct_call),
                    },
                    span.clone(),
                ));
            }
            // Skip the magic-`__call` wrap for receivers that are
            // already heavyweight expressions (Calls, Lambdas, Arrays,
            // etc.) — wrapping them again multiplies the AST depth
            // through the typeof check that re-traverses the receiver,
            // and chains of those would overflow the recursive
            // compiler walker. Simple identifiers, member accesses,
            // and previously-wrapped Sequences are cheap to clone
            // since their structure is already shallow.
            let recv_is_simple = matches!(
                &member_object.kind,
                ExprKind::Ident(_)
                    | ExprKind::Member { .. }
                    | ExprKind::This
                    | ExprKind::Sequence(_)
            );
            if method_name.starts_with("__")
                || matches!(&member_object.kind, ExprKind::This)
                || !recv_is_simple
            {
                let direct_member = Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(member_object),
                        field: method_name,
                        null_safe,
                    },
                    span.clone(),
                );
                return Ok(Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(direct_member),
                        args,
                        optional: null_safe,
                    },
                    span.clone(),
                ));
            }
            Ok(build_magic_call_rewrite(
                member_object,
                method_name,
                args,
                span,
            ))
        }
        Rule::property_access_op => {
            // Grammar: ("?->"|"->") ~ member_name. The arrow is a
            // literal token (pest does not yield it as a child pair),
            // so the only inner rule pair is `member_name`. Read
            // null_safe from the outer pair's source text.
            let null_safe = op.as_str().trim_start().starts_with("?->");
            let name_pair = op
                .into_inner()
                .next()
                .ok_or("property_access_op: missing name")?;
            let name = name_pair.into_inner().next().unwrap().as_str().to_string();
            let member = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(receiver),
                    field: name.clone(),
                    null_safe,
                },
                span.clone(),
            );
            // PHP `__get` magic method: when reading `$obj->prop` and
            // `prop` isn't an own property of `$obj`, dispatch through
            // `$obj->__get("prop")` if the class defines that magic
            // method. Walker wraps the read in:
            //
            //   ($_t = $obj,
            //    typeof $_t->prop !== "undefined"
            //      ? $_t->prop
            //      : (typeof $_t->__get === "function"
            //          ? $_t->__get("prop")
            //          : $_t->prop))
            //
            // Skipped for null-safe access (`?->` keeps its short-
            // circuit semantics) and for the OUTERMOST chain op when
            // we're walking an assignment LHS — that op is the WRITE
            // target and must remain a plain Member l-value. Skipped
            // for `__` prefixed names so internal accesses (`__type`,
            // `__call`, etc.) bypass the wrap and don't cause
            // infinite recursion via the `__get` lookup itself.
            if null_safe || is_assign_target || name.starts_with("__") {
                return Ok(member);
            }
            // Skip magic-get wrap when the receiver is `$this`. Inside
            // class methods the receiver is the instance — direct
            // property access is what user code expects, and the
            // wrap interferes with chained writes like
            // `$this->data[$k] = $v` (Index { Sequence(...), $k } as
            // assign target — the inner Sequence return value isn't
            // tracked as an l-value through the indexed write).
            if let ExprKind::Member { object, .. } = &member.kind {
                if is_php_this_expr(object) {
                    return Ok(member);
                }
            }
            // Extract receiver from member to use in temp save.
            let recv_for_save = match &member.kind {
                ExprKind::Member { object, .. } => (**object).clone(),
                _ => return Ok(member),
            };
            Ok(build_magic_get_rewrite(recv_for_save, name, span))
        }
        Rule::static_access_op => {
            // Grammar: `::` ~ class_member_name where class_member_name
            // can be `kw_class | identifier | variable | "{" expr "}"`.
            // For static fields PHP uses the `$variable` form
            // (`Class::$staticField`) — strip the leading `$` so the
            // member name matches the field key written by the
            // class-static-field initialiser. For `Class::class` the
            // member is the literal `"class"` (PHP class-name reflection).
            let mut inner = op.into_inner();
            let name_pair = inner.next().unwrap();
            let inner_pair = name_pair.into_inner().next().unwrap();
            let raw = inner_pair.as_str();
            let name = if matches!(inner_pair.as_rule(), Rule::variable) {
                raw.strip_prefix('$').unwrap_or(raw).to_string()
            } else {
                raw.to_string()
            };
            Ok(Expression::with_span(
                ExprKind::StaticAccess {
                    class: Box::new(receiver),
                    member: Box::new(Expression::ident(&name)),
                },
                span.clone(),
            ))
        }
        Rule::array_index_op => {
            let mut inner = op.into_inner();
            let index = if let Some(i) = inner.next() {
                walk_expression(i)?
            } else {
                Expression::null()
            };
            Ok(Expression::with_span(
                ExprKind::Index {
                    object: Box::new(receiver),
                    index: Box::new(index),
                    null_safe: false,
                },
                span.clone(),
            ))
        }
        Rule::call_op => {
            let mut inner = op.into_inner();
            // PHP 8.1 first-class callable: `strlen(...)` creates a
            // Closure bound to `strlen`. Rewrite to an arrow function
            // that forwards via rest-spread:
            //   strlen(...)  →  (...args) => strlen(...args)
            // This produces a Closure-typed value in the common AST so
            // assignment/passing/invoking all work without PHP-specific
            // handling downstream.
            let first = inner.next();
            if let Some(p) = &first {
                if matches!(p.as_rule(), Rule::first_class_callable_op) {
                    return Ok(php_first_class_callable_lambda(
                        php_first_class_callable_target(receiver, span),
                        false,
                        span,
                    ));
                }
            }
            let args = match first {
                Some(p) if matches!(p.as_rule(), Rule::arg_list) => walk_args(p)?,
                _ => Vec::new(),
            };
            // Normalize PHP-specific argument conventions to the common
            // AST's canonical order BEFORE the compiler sees them. Once
            // in the common AST, PHP and JS calls should be
            // indistinguishable — the compiler emits a single canonical
            // host call regardless of surface syntax.
            let args = canonicalize_php_call_args(&receiver, args);
            // PHP `Fiber::suspend($v)` → `__php_fiber_suspend($v)`
            // which emits the WASM `SUSPEND` op directly. Walker
            // strips the static-call shape so the rest of the
            // pipeline doesn't try to look up `Fiber.suspend` on a
            // class object that doesn't exist.
            if let ExprKind::StaticAccess { class, member } = &receiver.kind {
                if let (ExprKind::Ident(class_name), ExprKind::Ident(member_name)) =
                    (&class.kind, &member.kind)
                {
                    if class_name.trim_start_matches('\\') == "Fiber" && member_name == "suspend" {
                        return Ok(Expression::with_span(
                            ExprKind::Call {
                                callee: Box::new(Expression::with_span(
                                    ExprKind::Ident("__php_fiber_suspend".to_string()),
                                    span.clone(),
                                )),
                                args,
                                optional: false,
                            },
                            span.clone(),
                        ));
                    }
                }
            }
            // PHP DateTime / DateTimeImmutable static factory methods
            // route through the PHP datetime adapter layer, same as the
            // instance-method rewrites above.
            if let ExprKind::StaticAccess { class, member } = &receiver.kind {
                if let (ExprKind::Ident(class_name), ExprKind::Ident(member_name)) =
                    (&class.kind, &member.kind)
                {
                    if member_name == "createFromFormat" {
                        let target_fn = match class_name.trim_start_matches('\\') {
                            "DateTime" => Some("__php_dt_create_from_format"),
                            "DateTimeImmutable" => Some("__php_dt_imm_create_from_format"),
                            _ => None,
                        };
                        if let Some(fname) = target_fn {
                            return Ok(Expression::with_span(
                                ExprKind::Call {
                                    callee: Box::new(Expression::with_span(
                                        ExprKind::Ident(fname.to_string()),
                                        span.clone(),
                                    )),
                                    args,
                                    optional: false,
                                },
                                span.clone(),
                            ));
                        }
                    }
                    // PHP `Closure::bind($closure, $obj, $scope?)` — bind a
                    // closure to an object by rewriting `$this` inside the
                    // closure body to a captured temp holding the target
                    // object. This keeps the fix in the PHP frontend and
                    // avoids relying on method-style lambda binding in the
                    // shared runtime.
                    if class_name.trim_start_matches('\\') == "Closure"
                        && member_name == "bind"
                        && args.len() >= 2
                    {
                        if let ExprKind::Lambda {
                            params,
                            body,
                            is_async,
                            captures,
                        } = &args[0].value.kind
                        {
                            let bound_obj_name = format!(
                                "$__php_closure_bind_obj_{}_{}",
                                span.start_line, span.start_col,
                            );
                            let mut rebound_captures = captures.clone();
                            if !rebound_captures
                                .iter()
                                .any(|capture| capture == &bound_obj_name)
                            {
                                rebound_captures.push(bound_obj_name.clone());
                            }
                            let save_obj = Expression::with_span(
                                ExprKind::Assign {
                                    target: Box::new(Expression::with_span(
                                        ExprKind::Ident(bound_obj_name.clone()),
                                        span.clone(),
                                    )),
                                    value: Box::new(args[1].value.clone()),
                                },
                                span.clone(),
                            );
                            let rebound_lambda = Expression::with_span(
                                ExprKind::Lambda {
                                    params: params.clone(),
                                    body: bind_this_in_lambda_body(body, &bound_obj_name),
                                    is_async: *is_async,
                                    captures: rebound_captures,
                                },
                                span.clone(),
                            );
                            return Ok(Expression::with_span(
                                ExprKind::Sequence(vec![save_obj, rebound_lambda]),
                                span.clone(),
                            ));
                        }
                    }
                    // PHP `Closure::fromCallable($callable)` — produces a
                    // Closure forwarding to the named callable. Rewrite at
                    // walker-time to a 4-arg pass-through arrow function,
                    // mirroring the PHP 8.1 first-class callable rewrite
                    // (`fn(...)` at Rule::call_op above). Forms handled:
                    //   'name'              — bare function/builtin
                    //   [$obj, 'method']    — instance method
                    //   ['Class', 'method'] — static method
                    if class_name.trim_start_matches('\\') == "Closure"
                        && member_name == "fromCallable"
                        && args.len() == 1
                    {
                        let mk_param = |n: &str| Param {
                            name: n.to_string(),
                            type_hint: None,
                            default: Some(Expression::with_span(
                                ExprKind::Lit(Literal::Null),
                                span.clone(),
                            )),
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: true,
                            is_nullable: true,
                        };
                        let mk_arg_ident = |n: &str| {
                            Argument::positional(Expression::with_span(
                                ExprKind::Ident(n.to_string()),
                                span.clone(),
                            ))
                        };
                        let body_call_opt: Option<Expression> = match &args[0].value.kind {
                            ExprKind::Lit(Literal::Str(name)) => Some(Expression::with_span(
                                ExprKind::Call {
                                    callee: Box::new(Expression::with_span(
                                        ExprKind::Ident(name.clone()),
                                        span.clone(),
                                    )),
                                    args: vec![
                                        mk_arg_ident("a"),
                                        mk_arg_ident("b"),
                                        mk_arg_ident("c"),
                                        mk_arg_ident("d"),
                                    ],
                                    optional: false,
                                },
                                span.clone(),
                            )),
                            ExprKind::Array(elems) if elems.len() == 2 => {
                                let recv = &elems[0].value;
                                let method = match &elems[1].value.kind {
                                    ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
                                    _ => None,
                                };
                                method.map(|m| {
                                    let callee = match &recv.kind {
                                        ExprKind::Lit(Literal::Str(cls)) => Expression::with_span(
                                            ExprKind::StaticAccess {
                                                class: Box::new(Expression::with_span(
                                                    ExprKind::Ident(cls.clone()),
                                                    span.clone(),
                                                )),
                                                member: Box::new(Expression::with_span(
                                                    ExprKind::Ident(m.clone()),
                                                    span.clone(),
                                                )),
                                            },
                                            span.clone(),
                                        ),
                                        _ => Expression::with_span(
                                            ExprKind::Member {
                                                object: Box::new(recv.clone()),
                                                field: m,
                                                null_safe: false,
                                            },
                                            span.clone(),
                                        ),
                                    };
                                    Expression::with_span(
                                        ExprKind::Call {
                                            callee: Box::new(callee),
                                            args: vec![
                                                mk_arg_ident("a"),
                                                mk_arg_ident("b"),
                                                mk_arg_ident("c"),
                                                mk_arg_ident("d"),
                                            ],
                                            optional: false,
                                        },
                                        span.clone(),
                                    )
                                })
                            }
                            _ => None,
                        };
                        if let Some(body_call) = body_call_opt {
                            return Ok(Expression::with_span(
                                ExprKind::Lambda {
                                    params: vec![
                                        mk_param("a"),
                                        mk_param("b"),
                                        mk_param("c"),
                                        mk_param("d"),
                                    ],
                                    body: LambdaBody::Expr(Box::new(body_call)),
                                    is_async: false,
                                    captures: vec![],
                                },
                                span.clone(),
                            ));
                        }
                    }
                }
            }
            // PHP `parent::method(args)` — calls the parent's method
            // bound to current `$this`. Walker normalises to the
            // `super.method(args)` Member-call shape so the existing
            // super-method dispatch in compile_call (calls.rs lines
            // 173+) handles `$this` rebinding correctly.
            if let ExprKind::StaticAccess { class, member } = &receiver.kind {
                if matches!(class.kind, ExprKind::Super) {
                    if let ExprKind::Ident(method_name) = &member.kind {
                        // `parent::__construct(args)` is the parent
                        // CONSTRUCTOR, not a parent method — normalise to
                        // the bare `super(args)` call shape that the
                        // super-ctor dispatch in compile_call handles
                        // (constructors live in their own ctor chunks, so
                        // a `super.__construct` member lookup finds
                        // nothing).
                        if method_name == "__construct" {
                            return Ok(Expression::with_span(
                                ExprKind::Call {
                                    callee: Box::new(Expression::with_span(
                                        ExprKind::Super,
                                        span.clone(),
                                    )),
                                    args,
                                    optional: false,
                                },
                                span.clone(),
                            ));
                        }
                        let super_member = Expression::with_span(
                            ExprKind::Member {
                                object: Box::new(Expression::with_span(
                                    ExprKind::Super,
                                    span.clone(),
                                )),
                                field: method_name.clone(),
                                null_safe: false,
                            },
                            span.clone(),
                        );
                        return Ok(Expression::with_span(
                            ExprKind::Call {
                                callee: Box::new(super_member),
                                args,
                                optional: false,
                            },
                            span.clone(),
                        ));
                    }
                }
            }
            // PHP `Class::method(args)` (StaticAccess + Call) is
            // normalised to `Class.method(args)` Member-call shape by
            // the `__callStatic` magic-rewrite below — its `direct_call`
            // branch uses Member-shape so the static-method dispatch
            // in compile_call (calls.rs ~600) fires and pushes the
            // class object as `$this` slot 0. That makes
            // `static::X` (walked as `$this::X`) resolve correctly
            // through late static binding inside the method body.
            // Rewrite PHP function names whose JS equivalent already
            // exists (Math.trunc / parseInt / Member-method calls / etc).
            // After this, the AST contains no PHP-specific call shape.
            if let Some(kind) = rewrite_php_call_to_js(&receiver, &args, &span) {
                return Ok(Expression::with_span(kind, span.clone()));
            }
            // PHP `__invoke` magic method: when invoking a value held in a
            // variable (`$obj(args)`), PHP dispatches through `$obj->__invoke()`
            // if the value is a class instance with that method. Walker
            // wraps the call in a typeof-discriminated ternary so
            // function values still go through CALL_REF directly while
            // class instances get the magic-method dispatch.
            //
            //   $obj(a, b)
            //     →
            //   typeof $obj === "function" ? $obj(a, b) : $obj->__invoke(a, b)
            //
            // The wrapping only fires when the chain root was a `$variable`
            // primary (`from_variable`) — bare-identifier function calls
            // (`strlen($s)`) don't need the magic dispatch and would lose
            // optimisation if wrapped.
            // Skip the magic-`__invoke` wrap when:
            //   - args use spread / named / by-ref (variadic shapes
            //     don't fit a fixed-arity ternary)
            //   - any arg is itself a Call expression (the wrap
            //     duplicates args across both ternary branches; calls
            //     in args would double-evaluate AND blow AST depth
            //     when nested)
            //   - any arg is itself a Sequence (already a wrap)
            // Skipping these falls back to the regular Call path —
            // the wrap is opt-in for shallow `$obj(args)` patterns
            // where simple-bench tests use it.
            let has_unwrappable_args = args.iter().any(|a| {
                a.spread
                    || a.by_ref
                    || a.name.is_some()
                    || matches!(
                        &a.value.kind,
                        ExprKind::Call { .. }
                            | ExprKind::Sequence(_)
                            | ExprKind::Lambda { .. }
                            | ExprKind::FunctionExpr(_)
                            | ExprKind::New { .. }
                            | ExprKind::ClassExpr { .. }
                    )
            });
            if from_variable && !has_unwrappable_args {
                if let ExprKind::Ident(_) = &receiver.kind {
                    return Ok(build_magic_invoke_rewrite(receiver, args, &span));
                }
            }
            // PHP `__callStatic` magic method: when `Class::method(args)`
            // is invoked and the method isn't a function on the class
            // object, PHP dispatches to `Class::__callStatic("method",
            // [args])`. Wrap StaticAccess + Call with a typeof check
            // similar to the instance-method __call rewrite. Only fires
            // when the class side is a plain Ident (not Super, not
            // computed) — those use distinct dispatch paths.
            if let ExprKind::StaticAccess { class, member } = &receiver.kind {
                if let (ExprKind::Ident(_), ExprKind::Ident(method_name)) =
                    (&class.kind, &member.kind)
                {
                    let mname = method_name.clone();
                    let class_expr = (**class).clone();
                    return Ok(build_magic_call_static_rewrite(
                        class_expr, mname, args, &span,
                    ));
                }
            }
            Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(receiver),
                    args,
                    optional: false,
                },
                span.clone(),
            ))
        }
        Rule::inc_dec_op => {
            // PHP `++` / `--` aren't C-style "add 1" — PHP defines them
            // with Perl-style string-character carry ("aa"++ → "ab",
            // "az"++ → "ba", "zz"++ → "aaa") AND PHP-flavored numeric
            // coercion for non-string inputs. Normalise both at walker
            // time so the AST carries a language-neutral call to a
            // stdlib helper:
            //
            //   $x++   →   ($tmp = $x, $x = __php_increment($x), $tmp)
            //   $x--   →   ($tmp = $x, $x = __php_decrement($x), $tmp)
            //
            // The Sequence form returns the OLD value (PHP postfix
            // semantics) — required by `yield $n++` and similar
            // expression-level uses. The temp is unique per call site
            // (TMP_COUNTER) so nested post-increments like
            // `$a++ + $b++` don't clobber each other.
            //
            // Downstream compilers, consumers, and other language
            // walkers see a plain Sequence + call + assign — no
            // compiler-side `if profile.php_*` flag needed.
            let helper = if op.as_str() == "++" {
                "__php_increment"
            } else {
                "__php_decrement"
            };
            let callee = Expression::with_span(ExprKind::Ident(helper.to_string()), span.clone());
            let call = Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: vec![Argument::positional(receiver.clone())],
                    optional: false,
                },
                span.clone(),
            );
            let tmp = next_tmp_name("post_inc");
            let tmp_save = Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(Expression::with_span(
                        ExprKind::Ident(tmp.clone()),
                        span.clone(),
                    )),
                    value: Box::new(receiver.clone()),
                },
                span.clone(),
            );
            let assign = Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(receiver.clone()),
                    value: Box::new(call),
                },
                span.clone(),
            );
            let read_tmp = Expression::with_span(ExprKind::Ident(tmp), span.clone());
            Ok(Expression::with_span(
                ExprKind::Sequence(vec![tmp_save, assign, read_tmp]),
                span.clone(),
            ))
        }
        _ => Ok(receiver),
    }
}

fn walk_args(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if !matches!(p.as_rule(), Rule::argument) {
            continue;
        }
        // Detect spread/by_ref by inspecting the source slice — pest
        // doesn't capture the literal `...` / `&` as named rules.
        let raw = p.as_str();
        let spread = raw.trim_start().starts_with("...");
        let by_ref = raw.trim_start().starts_with('&');

        let mut name: Option<String> = None;
        let mut value: Option<Expression> = None;
        for sub in p.into_inner() {
            match sub.as_rule() {
                Rule::arg_name => name = Some(sub.as_str().to_string()),
                Rule::expression => value = Some(walk_expression(sub)?),
                _ => {}
            }
        }
        if let Some(v) = value {
            out.push(Argument {
                name,
                value: v,
                by_ref,
                spread,
            });
        }
    }
    Ok(out)
}

fn php_first_class_callable_target(receiver: Expression, span: &Span) -> Expression {
    match &receiver.kind {
        ExprKind::Ident(name) => {
            let mapped = match name.as_str() {
                "abs" => Some(("Math", "abs")),
                "round" => Some(("Math", "round")),
                "intval" => Some(("Number", "parseInt")),
                "ceil" => Some(("Math", "ceil")),
                "floor" => Some(("Math", "floor")),
                "sqrt" => Some(("Math", "sqrt")),
                "pow" => Some(("Math", "pow")),
                "exp" => Some(("Math", "exp")),
                "log" => Some(("Math", "log")),
                "log2" => Some(("Math", "log2")),
                "log10" => Some(("Math", "log10")),
                "sin" => Some(("Math", "sin")),
                "cos" => Some(("Math", "cos")),
                "tan" => Some(("Math", "tan")),
                "asin" => Some(("Math", "asin")),
                "acos" => Some(("Math", "acos")),
                "atan" => Some(("Math", "atan")),
                "atan2" => Some(("Math", "atan2")),
                "sinh" => Some(("Math", "sinh")),
                "cosh" => Some(("Math", "cosh")),
                "tanh" => Some(("Math", "tanh")),
                "asinh" => Some(("Math", "asinh")),
                "acosh" => Some(("Math", "acosh")),
                "atanh" => Some(("Math", "atanh")),
                _ => None,
            };
            if let Some((object, field)) = mapped {
                return Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(Expression::with_span(
                            ExprKind::Ident(object.to_string()),
                            span.clone(),
                        )),
                        field: field.to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                );
            }
            receiver
        }
        _ => receiver,
    }
}

fn php_callable_target_expr(expr: Expression, span: &Span) -> Expression {
    match expr.kind {
        ExprKind::Lit(Literal::Str(name)) => {
            Expression::with_span(ExprKind::Ident(name), span.clone())
        }
        ExprKind::Array(elements) if elements.len() == 2 => {
            let receiver = elements[0].value.clone();
            let member_name = match &elements[1].value.kind {
                ExprKind::Lit(Literal::Str(name)) => Some(name.clone()),
                _ => None,
            };
            if let Some(member_name) = member_name {
                match receiver.kind {
                    ExprKind::Lit(Literal::Str(class_name)) => Expression::with_span(
                        ExprKind::StaticAccess {
                            class: Box::new(Expression::with_span(
                                ExprKind::Ident(class_name),
                                span.clone(),
                            )),
                            member: Box::new(Expression::with_span(
                                ExprKind::Ident(member_name),
                                span.clone(),
                            )),
                        },
                        span.clone(),
                    ),
                    _ => Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(receiver),
                            field: member_name,
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                }
            } else {
                Expression::with_span(ExprKind::Array(elements), span.clone())
            }
        }
        _ => Expression::with_span(expr.kind, span.clone()),
    }
}

fn php_first_class_callable_lambda(callee: Expression, optional: bool, span: &Span) -> Expression {
    let direct_arity = match &callee.kind {
        ExprKind::Ident(name) => match name.as_str() {
            "strlen" | "strtoupper" | "strtolower" | "trim" | "ltrim" | "rtrim" | "strrev"
            | "strval" | "intval" => Some(1usize),
            _ => None,
        },
        ExprKind::Member { object, field, .. } => match (&object.kind, field.as_str()) {
            (ExprKind::Ident(obj), "abs") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "round") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "ceil") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "floor") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "sqrt") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "pow") if obj == "Math" => Some(2),
            (ExprKind::Ident(obj), "exp") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "log") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "log2") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "log10") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "sin") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "cos") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "tan") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "asin") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "acos") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "atan") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "atan2") if obj == "Math" => Some(2),
            (ExprKind::Ident(obj), "sinh") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "cosh") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "tanh") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "asinh") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "acosh") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "atanh") if obj == "Math" => Some(1),
            (ExprKind::Ident(obj), "parseInt") if obj == "Number" => Some(1),
            _ => None,
        },
        _ => None,
    };
    if let Some(direct_arity) = direct_arity {
        let params = ["a", "b", "c", "d"]
            .into_iter()
            .take(direct_arity)
            .map(|name| Param {
                name: name.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect::<Vec<_>>();
        let args = ["a", "b", "c", "d"]
            .into_iter()
            .take(direct_arity)
            .map(|name| {
                Argument::positional(Expression::with_span(
                    ExprKind::Ident(name.to_string()),
                    span.clone(),
                ))
            })
            .collect::<Vec<_>>();
        return Expression::with_span(
            ExprKind::Lambda {
                params,
                body: LambdaBody::Expr(Box::new(Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(callee),
                        args,
                        optional,
                    },
                    span.clone(),
                ))),
                is_async: false,
                captures: vec![],
            },
            span.clone(),
        );
    }

    let mk_param = |name: &str| Param {
        name: name.to_string(),
        type_hint: None,
        default: Some(Expression::with_span(
            ExprKind::Lit(Literal::Null),
            span.clone(),
        )),
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: true,
    };
    let mk_ident =
        |name: &str| Expression::with_span(ExprKind::Ident(name.to_string()), span.clone());
    let mk_arg = |name: &str| Argument::positional(mk_ident(name));
    let mk_null = || Expression::with_span(ExprKind::Lit(Literal::Null), span.clone());
    let target_ident = mk_ident("__fcc_target");
    let mk_is_null = |name: &str| {
        Expression::with_span(
            ExprKind::Binary {
                op: BinOp::StrictEq,
                left: Box::new(mk_ident(name)),
                right: Box::new(mk_null()),
            },
            span.clone(),
        )
    };
    let mk_call = |args: Vec<Argument>| {
        Expression::with_span(
            ExprKind::Call {
                callee: Box::new(target_ident.clone()),
                args,
                optional,
            },
            span.clone(),
        )
    };
    let ternary = |cond: Expression, then: Expression, else_: Expression| {
        Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(then),
                else_: Box::new(else_),
            },
            span.clone(),
        )
    };
    let body_call = ternary(
        mk_is_null("a"),
        mk_call(vec![]),
        ternary(
            mk_is_null("b"),
            mk_call(vec![mk_arg("a")]),
            ternary(
                mk_is_null("c"),
                mk_call(vec![mk_arg("a"), mk_arg("b")]),
                ternary(
                    mk_is_null("d"),
                    mk_call(vec![mk_arg("a"), mk_arg("b"), mk_arg("c")]),
                    mk_call(vec![mk_arg("a"), mk_arg("b"), mk_arg("c"), mk_arg("d")]),
                ),
            ),
        ),
    );
    let inner = Expression::with_span(
        ExprKind::Lambda {
            params: vec![mk_param("a"), mk_param("b"), mk_param("c"), mk_param("d")],
            body: LambdaBody::Expr(Box::new(body_call)),
            is_async: false,
            captures: vec!["__fcc_target".to_string()],
        },
        span.clone(),
    );
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Lambda {
                    params: vec![Param {
                        name: "__fcc_target".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    }],
                    body: LambdaBody::Expr(Box::new(inner)),
                    is_async: false,
                    captures: vec![],
                },
                span.clone(),
            )),
            args: vec![Argument::positional(callee)],
            optional: false,
        },
        span.clone(),
    )
}

fn walk_new(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    // new_expression = { kw_new ~ (anonymous_class | kw_static | kw_self | kw_parent
    //                              | qualified_name | variable | "(" expr ")")
    //                    ~ ("(" arg_list? ")")? }
    let mut class: Option<Expression> = None;
    let mut args: Vec<Argument> = Vec::new();
    // Iterate raw inner pairs (don't filter keywords) so `kw_static` /
    // `kw_self` / `kw_parent` as the class designator are visible. The
    // outer `kw_new` is the first child — skip it explicitly.
    let mut iter = pair.into_inner().peekable();
    if let Some(first) = iter.peek() {
        if matches!(first.as_rule(), Rule::kw_new) {
            iter.next();
        }
    }
    for p in iter {
        match p.as_rule() {
            // PHP 8 `new static(...)` / `new self(...)` / `new parent(...)`
            // — late-static-binding instantiation. Map each to the
            // existing context expressions so the rest of the
            // walker/compiler treats them identically:
            //   static → This (the class object passed as $this slot
            //                   in static method dispatch)
            //   self   → Ident(<current class name>) (set by walker
            //                   class-context push)
            //   parent → Super
            Rule::kw_static => {
                // In a STATIC method `$this` slot 0 holds the class object
                // itself (callable as ctor); in an INSTANCE method it holds
                // the instance, whose class is reachable through the
                // prototype-chain `constructor` link. Discriminate at
                // runtime so `new static(...)` works in both contexts:
                //   typeof $this === "function" ? $this : $this.constructor
                let this_e = Expression::with_span(ExprKind::This, span.clone());
                let typeof_this = Expression::with_span(
                    ExprKind::TypeOf(Box::new(this_e.clone())),
                    span.clone(),
                );
                let fn_str = Expression::with_span(
                    ExprKind::Lit(Literal::Str("function".to_string())),
                    span.clone(),
                );
                let is_fn = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(typeof_this),
                        right: Box::new(fn_str),
                    },
                    span.clone(),
                );
                let ctor_member = Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(this_e.clone()),
                        field: "constructor".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                );
                class = Some(Expression::with_span(
                    ExprKind::Ternary {
                        cond: Box::new(is_fn),
                        then: Box::new(this_e),
                        else_: Box::new(ctor_member),
                    },
                    span.clone(),
                ));
            }
            Rule::kw_self => {
                let cn = current_class_name().unwrap_or_default();
                class = Some(Expression::with_span(ExprKind::Ident(cn), span.clone()));
            }
            Rule::kw_parent => {
                class = Some(Expression::with_span(ExprKind::Super, span.clone()));
            }
            Rule::arg_list => args = walk_args(p)?,
            Rule::anonymous_class => {
                // PHP 8: `new class(args) extends Base implements I { ... }`.
                // Walk to a ClassExpr value; the surrounding `New` then
                // instantiates it with the provided ctor args.
                //
                // Inspect the source slice to know whether the first
                // qualified_name follows `extends` (parent class) or
                // `implements` (interface list). Pest doesn't yield
                // `kw_extends` / `kw_implements` as child pairs, so the
                // children are: optional arg_list, then qualified_names,
                // then class members. We check the source text to
                // determine whether the FIRST qualified name is a parent
                // or an interface.
                let raw = p.as_str();
                let has_extends = raw.contains(" extends ");
                let mut parent: Option<Expression> = None;
                let mut members: Vec<ClassMember> = Vec::new();
                let mut ctor_args: Vec<Argument> = Vec::new();
                let mut interfaces: Vec<String> = Vec::new();
                let mut first_qualified = true;
                for sub in inner_nokw(p) {
                    match sub.as_rule() {
                        Rule::arg_list
                            if ctor_args.is_empty() && parent.is_none() && members.is_empty() =>
                        {
                            ctor_args = walk_args(sub)?;
                        }
                        Rule::qualified_name => {
                            if first_qualified && has_extends {
                                parent = Some(Expression::ident(sub.as_str()));
                            } else {
                                interfaces.push(sub.as_str().to_string());
                            }
                            first_qualified = false;
                        }
                        Rule::use_trait
                        | Rule::class_constant
                        | Rule::property_declaration
                        | Rule::method_declaration
                        | Rule::empty_statement => {
                            if let Some(m) = walk_class_member(sub)? {
                                members.push(m);
                            }
                        }
                        _ => {}
                    }
                }
                let _ = interfaces; // walker doesn't enforce interface contracts
                args = ctor_args;
                class = Some(Expression::with_span(
                    ExprKind::ClassExpr {
                        name: None,
                        parent: parent.map(Box::new),
                        members,
                    },
                    span.clone(),
                ));
            }
            _ => {
                if class.is_none() {
                    class = Some(walk_expression(p)?);
                }
            }
        }
    }
    let class_expr = class.ok_or("new: missing class designator")?;
    // PHP `new Fiber($cb)` → `__php_fiber_new($cb)` which emits
    // `CONT_NEW` on the callback. Walker normalises so the rest of
    // the pipeline never sees `Fiber` as a class name; the
    // continuation Object that comes out is what `$fiber->start()`
    // and `Fiber::suspend()` operate on.
    if let ExprKind::Ident(class_name) = &class_expr.kind {
        if class_name.trim_start_matches('\\') == "Fiber" {
            return Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(Expression::with_span(
                        ExprKind::Ident("__php_fiber_new".to_string()),
                        span.clone(),
                    )),
                    args,
                    optional: false,
                },
                span.clone(),
            ));
        }
    }
    // PHP DateTime / DateTimeImmutable / DateInterval — rewrite to a
    // bare call against the bytecode adapter binding so the
    // `emitter/php/datetime_adapter.rs` emit_* functions handle the
    // construction. Avoids registering host fns and keeps the call
    // shape JS-uniform downstream.
    if let ExprKind::Ident(class_name) = &class_expr.kind {
        let rewrite_target: Option<&str> = match class_name.trim_start_matches('\\') {
            "DateTime" => Some("__php_dt_new"),
            "DateTimeImmutable" => Some("__php_dt_imm_new"),
            "DateInterval" => Some("__php_dateinterval_new"),
            _ => None,
        };
        if let Some(target) = rewrite_target {
            if target == "__php_dateinterval_new" {
                // DateInterval(P1Y2M3D) — for STRING-LITERAL ISO
                // arguments, parse at compile time and synthesize the
                // y/m/d/h/i/s components as numeric literals so the
                // adapter can emit them as constants. Dynamic strings
                // fall through to a runtime parser path (TODO).
                if let Some(arg) = args.first() {
                    if let ExprKind::Lit(Literal::Str(s)) = &arg.value.kind {
                        let (y, mo, d, h, mi, se) =
                            crate::emitter::php::datetime_adapter::parse_iso_duration(s);
                        return Ok(Expression::with_span(
                            ExprKind::Call {
                                callee: Box::new(Expression::with_span(
                                    ExprKind::Ident("__php_dateinterval_components".to_string()),
                                    span.clone(),
                                )),
                                args: vec![
                                    Argument::positional(Expression::int(y)),
                                    Argument::positional(Expression::int(mo)),
                                    Argument::positional(Expression::int(d)),
                                    Argument::positional(Expression::int(h)),
                                    Argument::positional(Expression::int(mi)),
                                    Argument::positional(Expression::int(se)),
                                ],
                                optional: false,
                            },
                            span,
                        ));
                    }
                }
            }
            return Ok(Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(Expression::with_span(
                        ExprKind::Ident(target.to_string()),
                        span.clone(),
                    )),
                    args,
                    optional: false,
                },
                span,
            ));
        }
    }
    // PHP exceptions use positional ctor args `(message, code, previous)`.
    // Normalize to the common JS-shaped `(message, {code, cause})` so the
    // shared exception emitter stamps `code`/`cause` onto the canonical
    // exception object. This keeps cross-language catch + `getPrevious()` /
    // `getCode()` working: a PHP-thrown exception's cause/code are then
    // visible to a JS/Python catcher, and vice-versa. `message` stays
    // positional; `code`/`previous` move into the options object.
    if let ExprKind::Ident(class_name) = &class_expr.kind {
        let bare = class_name.trim_start_matches('\\');
        if crate::emitter::errors::is_exception_type(bare)
            && !bare.eq_ignore_ascii_case("AggregateError")
            && args.len() >= 2
        {
            let msg = args[0].clone();
            let mut props: Vec<ObjectProperty> = vec![ObjectProperty::KeyValue {
                key: Expression::string("code"),
                value: args[1].value.clone(),
            }];
            if let Some(prev) = args.get(2) {
                props.push(ObjectProperty::KeyValue {
                    key: Expression::string("cause"),
                    value: prev.value.clone(),
                });
            }
            let opts = Expression::with_span(ExprKind::Object(props), span.clone());
            return Ok(Expression::with_span(
                ExprKind::New {
                    class: Box::new(class_expr),
                    args: vec![msg, Argument::positional(opts)],
                },
                span,
            ));
        }
    }
    Ok(Expression::with_span(
        ExprKind::New {
            class: Box::new(class_expr),
            args,
        },
        span,
    ))
}

fn walk_match(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = inner_nokw(pair);
    let subject = walk_expression(inner.next().unwrap())?;
    let mut arms: Vec<MatchArm> = Vec::new();
    for p in inner {
        if !matches!(p.as_rule(), Rule::match_arm) {
            continue;
        }
        // match_arm = { (kw_default | match_conditions) ~ "=>" ~ expression }
        // Use the source slice to detect default-arms because kw_default
        // is filtered out alongside other keyword tokens.
        let arm_src = p.as_str().trim_start();
        let is_default = arm_src.to_lowercase().starts_with("default");
        let mut conditions: Option<Vec<Expression>> = None;
        let mut body: Option<Expression> = None;
        for sub in inner_nokw(p) {
            match sub.as_rule() {
                Rule::match_conditions => {
                    let exprs: Result<Vec<_>, _> = sub.into_inner().map(walk_expression).collect();
                    conditions = Some(exprs?);
                }
                Rule::expression => body = Some(walk_expression(sub)?),
                _ => {}
            }
        }
        if is_default {
            conditions = None;
        }
        arms.push(MatchArm {
            conditions,
            body: body.unwrap_or_else(Expression::null),
        });
    }
    // PHP `match` throws UnhandledMatchError when no arm matches AND no
    // default is given. Synthesize a default that throws so the runtime
    // semantics match — JS/Vybe `Match` falls through to null otherwise.
    let has_default = arms.iter().any(|a| a.conditions.is_none());
    if !has_default {
        let throw_expr = Expression::new(ExprKind::New {
            class: Box::new(Expression::ident("UnhandledMatchError")),
            args: vec![Argument::positional(Expression::string(
                "Unhandled match value",
            ))],
        });
        let throw_stmt = Statement::new(StmtKind::Throw {
            expr: Some(throw_expr),
            cause: None,
        });
        let throw_lambda = Expression::with_span(
            ExprKind::Call {
                callee: Box::new(Expression::with_span(
                    ExprKind::Lambda {
                        params: vec![],
                        body: LambdaBody::Block(vec![throw_stmt]),
                        is_async: false,
                        captures: vec![],
                    },
                    span.clone(),
                )),
                args: vec![],
                optional: false,
            },
            span.clone(),
        );
        arms.push(MatchArm {
            conditions: None,
            body: throw_lambda,
        });
    }
    Ok(Expression::with_span(
        ExprKind::Match {
            subject: Box::new(subject),
            arms,
        },
        span,
    ))
}

fn walk_array(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let is_short_array = matches!(pair.as_rule(), Rule::short_array_expression);
    let source_start = pair.as_span().start();
    let mut elems = Vec::new();
    let mut previous_end = source_start + 1;
    for p in pair.into_inner() {
        if !matches!(p.as_rule(), Rule::array_element) {
            continue;
        }
        if is_short_array {
            let gap = &p.as_span().get_input()[previous_end..p.as_span().start()];
            let comma_count = gap.chars().filter(|ch| *ch == ',').count();
            let holes = if elems.is_empty() {
                comma_count
            } else {
                comma_count.saturating_sub(1)
            };
            for _ in 0..holes {
                elems.push(ArrayElement {
                    key: None,
                    value: Expression::null(),
                    spread: false,
                    by_ref: false,
                });
            }
        }
        previous_end = p.as_span().end();
        // The grammar's leading `"..."` literal produces no pair, so detect
        // a spread element (`[...$gen]`, `[...$arr]`) from the source text.
        let is_spread = p.as_str().trim_start().starts_with("...");
        let mut sub_iter = p.into_inner();
        let first = sub_iter.next();
        let second = sub_iter.next();
        match (first, second) {
            (Some(first), Some(second)) => {
                // key => value
                let key = walk_expression(first)?;
                let value = walk_expression(second)?;
                elems.push(ArrayElement {
                    key: Some(key),
                    value,
                    spread: false,
                    by_ref: false,
                });
            }
            (Some(first), None) => {
                let value = walk_expression(first)?;
                elems.push(ArrayElement {
                    key: None,
                    value,
                    spread: is_spread,
                    by_ref: false,
                });
            }
            _ => {}
        }
    }
    Ok(Expression::with_span(ExprKind::Array(elems), span))
}

fn walk_foreach_value_target(pair: Pair<Rule>) -> Result<Expression, String> {
    Ok(expression_into_destructure_target(walk_expression(pair)?))
}

fn foreach_binding_target(
    target: Expression,
    suffix: usize,
) -> Result<(String, Option<Statement>), String> {
    match target.kind {
        ExprKind::Ident(name) => Ok((name, None)),
        ExprKind::Destructure(pattern) => {
            let tmp = format!("__php_foreach_item_{}", suffix);
            let assign = Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(Expression::with_span(
                        ExprKind::Destructure(pattern),
                        target.span,
                    )),
                    value: Box::new(Expression::with_span(
                        ExprKind::Ident(tmp.clone()),
                        target.span,
                    )),
                },
                target.span,
            );
            Ok((
                tmp,
                Some(Statement::with_span(StmtKind::Expr(assign), target.span)),
            ))
        }
        _ => Err("foreach: unsupported target".into()),
    }
}

fn expression_into_destructure_target(expr: Expression) -> Expression {
    if let Some(pattern) = expression_to_destructure_pattern(&expr) {
        Expression::with_span(ExprKind::Destructure(pattern), expr.span)
    } else {
        expr
    }
}

fn expression_to_destructure_pattern(expr: &Expression) -> Option<DestructurePattern> {
    match &expr.kind {
        ExprKind::Destructure(pattern) => Some(pattern.clone()),
        ExprKind::Array(elems) => array_elements_to_destructure_pattern(elems),
        _ => None,
    }
}

fn array_elements_to_destructure_pattern(elems: &[ArrayElement]) -> Option<DestructurePattern> {
    let has_keys = elems.iter().any(|elem| elem.key.is_some());
    if has_keys {
        let mut props = Vec::with_capacity(elems.len());
        for elem in elems {
            let key_expr = elem.key.as_ref()?;
            let key = literal_key_name(key_expr)?;
            let value = expression_to_binding_pattern(&elem.value)?;
            props.push(ObjectPatternProp {
                key,
                value: Some(value),
                default: None,
                is_rest: false,
            });
        }
        Some(DestructurePattern::Object(props))
    } else {
        let mut out = Vec::with_capacity(elems.len());
        for elem in elems {
            if elem.spread {
                if let ExprKind::Ident(name) = &elem.value.kind {
                    out.push(ArrayPatternElem::Rest(name.clone()));
                    continue;
                }
            }
            if let Some(pat) = expression_to_binding_pattern(&elem.value) {
                out.push(ArrayPatternElem::Pattern(pat, None));
            } else {
                out.push(ArrayPatternElem::Hole);
            }
        }
        Some(DestructurePattern::Array(out))
    }
}

fn expression_to_binding_pattern(expr: &Expression) -> Option<BindingPattern> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(BindingPattern::Ident(name.clone())),
        ExprKind::Destructure(DestructurePattern::Array(elems)) => {
            Some(BindingPattern::Array(elems.clone()))
        }
        ExprKind::Destructure(DestructurePattern::Object(props)) => {
            Some(BindingPattern::Object(props.clone()))
        }
        ExprKind::Array(elems) => match array_elements_to_destructure_pattern(elems)? {
            DestructurePattern::Array(elems) => Some(BindingPattern::Array(elems)),
            DestructurePattern::Object(props) => Some(BindingPattern::Object(props)),
        },
        _ => None,
    }
}

fn literal_key_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
        ExprKind::Lit(Literal::Int(n)) => Some(n.to_string()),
        _ => None,
    }
}

fn is_php_this_expr(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::This)
        || matches!(&expr.kind, ExprKind::Ident(name) if name == "$this")
}

fn walk_closure(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut params: Vec<Param> = Vec::new();
    let mut captures: Vec<String> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::closure_use => {
                for v in p.into_inner() {
                    if matches!(v.as_rule(), Rule::closure_use_var) {
                        let by_ref = v.as_str().trim_start().starts_with('&');
                        if let Some(var) = v
                            .into_inner()
                            .find(|q| matches!(q.as_rule(), Rule::variable))
                        {
                            let capture_name = strip_dollar(var.as_str()).to_string();
                            if by_ref {
                                captures.push(format!("&{capture_name}"));
                            } else {
                                captures.push(capture_name);
                            }
                        }
                    }
                }
            }
            Rule::block_statement => body = walk_statement_into_body(p)?,
            _ => {}
        }
    }
    body = lower_php_runtime_arg_helpers_in_block(&mut params, body);
    Ok(Expression::with_span(
        ExprKind::Lambda {
            params,
            body: LambdaBody::Block(body),
            is_async: false,
            captures,
        },
        span,
    ))
}

fn walk_arrow_function(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut params: Vec<Param> = Vec::new();
    let mut body_expr: Option<Expression> = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::expression => body_expr = Some(walk_expression(p)?),
            _ => {}
        }
    }
    let body_expr = lower_php_runtime_arg_helpers_in_expr(
        &mut params,
        body_expr.unwrap_or_else(Expression::null),
    );
    Ok(Expression::with_span(
        ExprKind::Lambda {
            params,
            body: LambdaBody::Expr(Box::new(body_expr)),
            is_async: false,
            captures: Vec::new(),
        },
        span,
    ))
}

const PHP_RUNTIME_ARGS_REST_NAME: &str = "__vybe_php_runtime_args";

fn php_runtime_args_rest_param() -> Param {
    Param {
        name: PHP_RUNTIME_ARGS_REST_NAME.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: true,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

fn ensure_php_runtime_args_rest_param(params: &mut Vec<Param>) {
    if !params.iter().any(|param| param.is_rest) {
        params.push(php_runtime_args_rest_param());
    }
}

fn php_runtime_args_array_expr(params: &[Param], span: &Span) -> Expression {
    Expression::with_span(
        ExprKind::Array(
            params
                .iter()
                .map(|param| ArrayElement {
                    key: None,
                    value: Expression::with_span(ExprKind::Ident(param.name.clone()), span.clone()),
                    spread: param.is_rest,
                    by_ref: false,
                })
                .collect(),
        ),
        span.clone(),
    )
}

fn php_runtime_arg_helper_name_and_arg_count(expr: &Expression) -> Option<(&str, usize)> {
    let ExprKind::Call {
        callee,
        args,
        optional: false,
    } = &expr.kind
    else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    match name.as_str() {
        "func_get_args" | "func_num_args" => Some((name.as_str(), args.len())),
        "func_get_arg" => Some((name.as_str(), args.len())),
        _ => None,
    }
}

fn php_expr_uses_runtime_arg_helpers(expr: &Expression) -> bool {
    if let Some((name, argc)) = php_runtime_arg_helper_name_and_arg_count(expr) {
        if matches!(
            (name, argc),
            ("func_get_args", 0) | ("func_num_args", 0) | ("func_get_arg", 1)
        ) {
            return true;
        }
    }

    match &expr.kind {
        ExprKind::Binary { left, right, .. } => {
            php_expr_uses_runtime_arg_helpers(left) || php_expr_uses_runtime_arg_helpers(right)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr) => php_expr_uses_runtime_arg_helpers(expr),
        ExprKind::Yield(expr) => expr
            .as_ref()
            .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr)),
        ExprKind::Ternary { cond, then, else_ } => {
            php_expr_uses_runtime_arg_helpers(cond)
                || php_expr_uses_runtime_arg_helpers(then)
                || php_expr_uses_runtime_arg_helpers(else_)
        }
        ExprKind::Member { object, .. } => php_expr_uses_runtime_arg_helpers(object),
        ExprKind::Index { object, index, .. } => {
            php_expr_uses_runtime_arg_helpers(object) || php_expr_uses_runtime_arg_helpers(index)
        }
        ExprKind::Call { callee, args, .. }
        | ExprKind::New {
            class: callee,
            args,
        } => {
            php_expr_uses_runtime_arg_helpers(callee)
                || args
                    .iter()
                    .any(|arg| php_expr_uses_runtime_arg_helpers(&arg.value))
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            php_expr_uses_runtime_arg_helpers(target) || php_expr_uses_runtime_arg_helpers(value)
        }
        ExprKind::Array(elements) => elements.iter().any(|element| {
            element
                .key
                .as_ref()
                .is_some_and(|key| php_expr_uses_runtime_arg_helpers(key))
                || php_expr_uses_runtime_arg_helpers(&element.value)
        }),
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            items.iter().any(php_expr_uses_runtime_arg_helpers)
        }
        ExprKind::Object(props) => props.iter().any(|prop| match prop {
            ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => {
                php_expr_uses_runtime_arg_helpers(key) || php_expr_uses_runtime_arg_helpers(value)
            }
            ObjectProperty::Spread(expr) => php_expr_uses_runtime_arg_helpers(expr),
            _ => false,
        }),
        ExprKind::Interpolation(parts) => parts.iter().any(|part| match part {
            InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                php_expr_uses_runtime_arg_helpers(expr)
            }
            _ => false,
        }),
        ExprKind::IsType { expr, .. } | ExprKind::Cast { expr, .. } => {
            php_expr_uses_runtime_arg_helpers(expr)
        }
        ExprKind::NullCoalesce { left, right } => {
            php_expr_uses_runtime_arg_helpers(left) || php_expr_uses_runtime_arg_helpers(right)
        }
        ExprKind::Comprehension { element, .. } => php_expr_uses_runtime_arg_helpers(element),
        ExprKind::Slice { lower, upper, step } => {
            lower
                .as_ref()
                .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr))
                || upper
                    .as_ref()
                    .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr))
                || step
                    .as_ref()
                    .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr))
        }
        ExprKind::Range { start, end, .. } => {
            php_expr_uses_runtime_arg_helpers(start) || php_expr_uses_runtime_arg_helpers(end)
        }
        ExprKind::StaticAccess { class, member } => {
            php_expr_uses_runtime_arg_helpers(class) || php_expr_uses_runtime_arg_helpers(member)
        }
        ExprKind::Match { subject, arms } => {
            php_expr_uses_runtime_arg_helpers(subject)
                || arms.iter().any(|arm| {
                    arm.conditions.as_ref().is_some_and(|conditions| {
                        conditions.iter().any(php_expr_uses_runtime_arg_helpers)
                    }) || php_expr_uses_runtime_arg_helpers(&arm.body)
                })
        }
        ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) | ExprKind::ClassExpr { .. } => false,
        _ => false,
    }
}

fn php_stmt_uses_runtime_arg_helpers(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) => php_expr_uses_runtime_arg_helpers(expr),
        StmtKind::Block(body) => body.iter().any(php_stmt_uses_runtime_arg_helpers),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            php_expr_uses_runtime_arg_helpers(cond)
                || then_body.iter().any(php_stmt_uses_runtime_arg_helpers)
                || elifs.iter().any(|(cond, body)| {
                    php_expr_uses_runtime_arg_helpers(cond)
                        || body.iter().any(php_stmt_uses_runtime_arg_helpers)
                })
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(php_stmt_uses_runtime_arg_helpers))
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            init.as_ref()
                .is_some_and(|stmt| php_stmt_uses_runtime_arg_helpers(stmt))
                || cond
                    .as_ref()
                    .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr))
                || update
                    .as_ref()
                    .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr))
                || body.iter().any(php_stmt_uses_runtime_arg_helpers)
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            php_expr_uses_runtime_arg_helpers(iter)
                || body.iter().any(php_stmt_uses_runtime_arg_helpers)
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(php_stmt_uses_runtime_arg_helpers))
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            php_expr_uses_runtime_arg_helpers(cond)
                || body.iter().any(php_stmt_uses_runtime_arg_helpers)
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(php_stmt_uses_runtime_arg_helpers))
        }
        StmtKind::DoWhile { body, cond, .. } => {
            body.iter().any(php_stmt_uses_runtime_arg_helpers)
                || php_expr_uses_runtime_arg_helpers(cond)
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            php_expr_uses_runtime_arg_helpers(expr)
                || cases.iter().any(|case| {
                    case.conditions.iter().any(|condition| match condition {
                        CaseCondition::Value(expr) => php_expr_uses_runtime_arg_helpers(expr),
                        CaseCondition::Range { from, to } => {
                            php_expr_uses_runtime_arg_helpers(from)
                                || php_expr_uses_runtime_arg_helpers(to)
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            php_expr_uses_runtime_arg_helpers(expr)
                        }
                    }) || case.body.iter().any(php_stmt_uses_runtime_arg_helpers)
                })
                || default
                    .as_ref()
                    .is_some_and(|body| body.iter().any(php_stmt_uses_runtime_arg_helpers))
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            body.iter().any(php_stmt_uses_runtime_arg_helpers)
                || catches.iter().any(|catch| {
                    catch.body.iter().any(php_stmt_uses_runtime_arg_helpers)
                        || catch
                            .when_clause
                            .as_ref()
                            .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr))
                })
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(php_stmt_uses_runtime_arg_helpers))
                || finally
                    .as_ref()
                    .is_some_and(|body| body.iter().any(php_stmt_uses_runtime_arg_helpers))
        }
        StmtKind::With { items, body, .. } => {
            items
                .iter()
                .any(|item| php_expr_uses_runtime_arg_helpers(&item.expr))
                || body.iter().any(php_stmt_uses_runtime_arg_helpers)
        }
        StmtKind::Using { resource, body, .. } => {
            php_expr_uses_runtime_arg_helpers(resource)
                || body.iter().any(php_stmt_uses_runtime_arg_helpers)
        }
        StmtKind::Lock { expr, body } => {
            php_expr_uses_runtime_arg_helpers(expr)
                || body.iter().any(php_stmt_uses_runtime_arg_helpers)
        }
        StmtKind::Return(expr) => expr
            .as_ref()
            .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr)),
        StmtKind::Throw { expr, cause } => {
            expr.as_ref()
                .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr))
                || cause
                    .as_ref()
                    .is_some_and(|expr| php_expr_uses_runtime_arg_helpers(expr))
        }
        StmtKind::Assign { targets, value } => {
            targets.iter().any(php_expr_uses_runtime_arg_helpers)
                || php_expr_uses_runtime_arg_helpers(value)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            php_expr_uses_runtime_arg_helpers(target) || php_expr_uses_runtime_arg_helpers(value)
        }
        StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => false,
        _ => false,
    }
}

fn lower_php_runtime_arg_helpers_in_block(
    params: &mut Vec<Param>,
    body: Vec<Statement>,
) -> Vec<Statement> {
    if !body.iter().any(php_stmt_uses_runtime_arg_helpers) {
        return body;
    }

    ensure_php_runtime_args_rest_param(params);
    body.iter()
        .map(|stmt| rewrite_php_runtime_arg_helpers_in_stmt(stmt, params))
        .collect()
}

fn lower_php_runtime_arg_helpers_in_expr(params: &mut Vec<Param>, expr: Expression) -> Expression {
    if !php_expr_uses_runtime_arg_helpers(&expr) {
        return expr;
    }

    ensure_php_runtime_args_rest_param(params);
    rewrite_php_runtime_arg_helpers_in_expr(&expr, params)
}

fn rewrite_php_runtime_arg_helpers_in_stmt(stmt: &Statement, params: &[Param]) -> Statement {
    let kind = match &stmt.kind {
        StmtKind::Expr(expr) => {
            StmtKind::Expr(rewrite_php_runtime_arg_helpers_in_expr(expr, params))
        }
        StmtKind::Block(body) => StmtKind::Block(
            body.iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
        ),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => StmtKind::If {
            cond: rewrite_php_runtime_arg_helpers_in_expr(cond, params),
            then_body: then_body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
            elifs: elifs
                .iter()
                .map(|(cond, body)| {
                    (
                        rewrite_php_runtime_arg_helpers_in_expr(cond, params),
                        body.iter()
                            .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                            .collect(),
                    )
                })
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                    .collect()
            }),
        },
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => StmtKind::For {
            init: init
                .as_ref()
                .map(|inner| Box::new(rewrite_php_runtime_arg_helpers_in_stmt(inner, params))),
            cond: cond
                .as_ref()
                .map(|expr| rewrite_php_runtime_arg_helpers_in_expr(expr, params)),
            update: update
                .as_ref()
                .map(|expr| rewrite_php_runtime_arg_helpers_in_expr(expr, params)),
            body: body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
        },
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => StmtKind::ForIn {
            var: var.clone(),
            key: key.clone(),
            iter: rewrite_php_runtime_arg_helpers_in_expr(iter, params),
            body: body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
            of: *of,
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                    .collect()
            }),
            is_async: *is_async,
        },
        StmtKind::While {
            cond,
            body,
            else_body,
        } => StmtKind::While {
            cond: rewrite_php_runtime_arg_helpers_in_expr(cond, params),
            body: body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                    .collect()
            }),
        },
        StmtKind::DoWhile { body, cond, until } => StmtKind::DoWhile {
            body: body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
            cond: rewrite_php_runtime_arg_helpers_in_expr(cond, params),
            until: *until,
        },
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => StmtKind::Switch {
            expr: rewrite_php_runtime_arg_helpers_in_expr(expr, params),
            cases: cases
                .iter()
                .map(|case| SwitchCase {
                    conditions: case
                        .conditions
                        .iter()
                        .map(|condition| match condition {
                            CaseCondition::Value(expr) => CaseCondition::Value(
                                rewrite_php_runtime_arg_helpers_in_expr(expr, params),
                            ),
                            CaseCondition::Range { from, to } => CaseCondition::Range {
                                from: rewrite_php_runtime_arg_helpers_in_expr(from, params),
                                to: rewrite_php_runtime_arg_helpers_in_expr(to, params),
                            },
                            CaseCondition::Comparison { op, expr } => CaseCondition::Comparison {
                                op: *op,
                                expr: rewrite_php_runtime_arg_helpers_in_expr(expr, params),
                            },
                        })
                        .collect(),
                    body: case
                        .body
                        .iter()
                        .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                        .collect(),
                })
                .collect(),
            default: default.as_ref().map(|body| {
                body.iter()
                    .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                    .collect()
            }),
        },
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => StmtKind::Try {
            body: body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
            catches: catches
                .iter()
                .map(|catch| CatchClause {
                    types: catch.types.clone(),
                    var_name: catch.var_name.clone(),
                    stack_var: catch.stack_var.clone(),
                    body: catch
                        .body
                        .iter()
                        .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                        .collect(),
                    when_clause: catch
                        .when_clause
                        .as_ref()
                        .map(|expr| rewrite_php_runtime_arg_helpers_in_expr(expr, params)),
                })
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                    .collect()
            }),
            finally: finally.as_ref().map(|body| {
                body.iter()
                    .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                    .collect()
            }),
        },
        StmtKind::With {
            items,
            body,
            is_async,
        } => StmtKind::With {
            items: items
                .iter()
                .map(|item| WithItem {
                    expr: rewrite_php_runtime_arg_helpers_in_expr(&item.expr, params),
                    var: item.var.clone(),
                })
                .collect(),
            body: body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
            is_async: *is_async,
        },
        StmtKind::Using {
            var,
            resource,
            body,
        } => StmtKind::Using {
            var: var.clone(),
            resource: rewrite_php_runtime_arg_helpers_in_expr(resource, params),
            body: body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
        },
        StmtKind::Lock { expr, body } => StmtKind::Lock {
            expr: rewrite_php_runtime_arg_helpers_in_expr(expr, params),
            body: body
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_stmt(inner, params))
                .collect(),
        },
        StmtKind::Return(expr) => StmtKind::Return(
            expr.as_ref()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_expr(inner, params)),
        ),
        StmtKind::Throw { expr, cause } => StmtKind::Throw {
            expr: expr
                .as_ref()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_expr(inner, params)),
            cause: cause
                .as_ref()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_expr(inner, params)),
        },
        StmtKind::Assign { targets, value } => StmtKind::Assign {
            targets: targets
                .iter()
                .map(|target| rewrite_php_runtime_arg_helpers_in_expr(target, params))
                .collect(),
            value: rewrite_php_runtime_arg_helpers_in_expr(value, params),
        },
        StmtKind::CompoundAssign { target, op, value } => StmtKind::CompoundAssign {
            target: rewrite_php_runtime_arg_helpers_in_expr(target, params),
            op: *op,
            value: rewrite_php_runtime_arg_helpers_in_expr(value, params),
        },
        _ => stmt.kind.clone(),
    };
    Statement::with_span(kind, stmt.span)
}

fn rewrite_php_runtime_arg_helpers_in_expr(expr: &Expression, params: &[Param]) -> Expression {
    let span = expr.span;
    if let Some((name, argc)) = php_runtime_arg_helper_name_and_arg_count(expr) {
        let runtime_args = php_runtime_args_array_expr(params, &span);
        match (name, argc) {
            ("func_get_args", 0) => return runtime_args,
            ("func_num_args", 0) => {
                return Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(Expression::with_span(
                            ExprKind::Ident("count".to_string()),
                            span,
                        )),
                        args: vec![Argument::positional(runtime_args)],
                        optional: false,
                    },
                    span,
                );
            }
            ("func_get_arg", 1) => {
                let ExprKind::Call { args, .. } = &expr.kind else {
                    unreachable!()
                };
                return Expression::with_span(
                    ExprKind::Index {
                        object: Box::new(runtime_args),
                        index: Box::new(rewrite_php_runtime_arg_helpers_in_expr(
                            &args[0].value,
                            params,
                        )),
                        null_safe: false,
                    },
                    span,
                );
            }
            _ => {}
        }
    }

    let kind = match &expr.kind {
        ExprKind::Binary { op, left, right } => ExprKind::Binary {
            op: *op,
            left: Box::new(rewrite_php_runtime_arg_helpers_in_expr(left, params)),
            right: Box::new(rewrite_php_runtime_arg_helpers_in_expr(right, params)),
        },
        ExprKind::Unary { op, expr } => ExprKind::Unary {
            op: *op,
            expr: Box::new(rewrite_php_runtime_arg_helpers_in_expr(expr, params)),
        },
        ExprKind::Ternary { cond, then, else_ } => ExprKind::Ternary {
            cond: Box::new(rewrite_php_runtime_arg_helpers_in_expr(cond, params)),
            then: Box::new(rewrite_php_runtime_arg_helpers_in_expr(then, params)),
            else_: Box::new(rewrite_php_runtime_arg_helpers_in_expr(else_, params)),
        },
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => ExprKind::Member {
            object: Box::new(rewrite_php_runtime_arg_helpers_in_expr(object, params)),
            field: field.clone(),
            null_safe: *null_safe,
        },
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => ExprKind::Index {
            object: Box::new(rewrite_php_runtime_arg_helpers_in_expr(object, params)),
            index: Box::new(rewrite_php_runtime_arg_helpers_in_expr(index, params)),
            null_safe: *null_safe,
        },
        ExprKind::Call {
            callee,
            args,
            optional,
        } => ExprKind::Call {
            callee: Box::new(rewrite_php_runtime_arg_helpers_in_expr(callee, params)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: rewrite_php_runtime_arg_helpers_in_expr(&arg.value, params),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        },
        ExprKind::New { class, args } => ExprKind::New {
            class: Box::new(rewrite_php_runtime_arg_helpers_in_expr(class, params)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: rewrite_php_runtime_arg_helpers_in_expr(&arg.value, params),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
        },
        ExprKind::Assign { target, value } => ExprKind::Assign {
            target: Box::new(rewrite_php_runtime_arg_helpers_in_expr(target, params)),
            value: Box::new(rewrite_php_runtime_arg_helpers_in_expr(value, params)),
        },
        ExprKind::Array(elements) => ExprKind::Array(
            elements
                .iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .as_ref()
                        .map(|key| rewrite_php_runtime_arg_helpers_in_expr(key, params)),
                    value: rewrite_php_runtime_arg_helpers_in_expr(&element.value, params),
                    spread: element.spread,
                    by_ref: element.by_ref,
                })
                .collect(),
        ),
        ExprKind::Tuple(items) => ExprKind::Tuple(
            items
                .iter()
                .map(|item| rewrite_php_runtime_arg_helpers_in_expr(item, params))
                .collect(),
        ),
        ExprKind::Set(items) => ExprKind::Set(
            items
                .iter()
                .map(|item| rewrite_php_runtime_arg_helpers_in_expr(item, params))
                .collect(),
        ),
        ExprKind::Object(props) => ExprKind::Object(
            props
                .iter()
                .map(|prop| match prop {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: rewrite_php_runtime_arg_helpers_in_expr(key, params),
                        value: rewrite_php_runtime_arg_helpers_in_expr(value, params),
                    },
                    ObjectProperty::Spread(expr) => ObjectProperty::Spread(
                        rewrite_php_runtime_arg_helpers_in_expr(expr, params),
                    ),
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: rewrite_php_runtime_arg_helpers_in_expr(key, params),
                        value: rewrite_php_runtime_arg_helpers_in_expr(value, params),
                    },
                    other => other.clone(),
                })
                .collect(),
        ),
        ExprKind::Interpolation(parts) => ExprKind::Interpolation(
            parts
                .iter()
                .map(|part| match part {
                    InterpolPart::Expr(expr) => {
                        InterpolPart::Expr(rewrite_php_runtime_arg_helpers_in_expr(expr, params))
                    }
                    InterpolPart::Formatted(expr, fmt) => InterpolPart::Formatted(
                        rewrite_php_runtime_arg_helpers_in_expr(expr, params),
                        fmt.clone(),
                    ),
                    other => other.clone(),
                })
                .collect(),
        ),
        ExprKind::IsType { expr, type_name } => ExprKind::IsType {
            expr: Box::new(rewrite_php_runtime_arg_helpers_in_expr(expr, params)),
            type_name: type_name.clone(),
        },
        ExprKind::Cast { expr, type_name } => ExprKind::Cast {
            expr: Box::new(rewrite_php_runtime_arg_helpers_in_expr(expr, params)),
            type_name: type_name.clone(),
        },
        ExprKind::TypeOf(expr) => ExprKind::TypeOf(Box::new(
            rewrite_php_runtime_arg_helpers_in_expr(expr, params),
        )),
        ExprKind::NullCoalesce { left, right } => ExprKind::NullCoalesce {
            left: Box::new(rewrite_php_runtime_arg_helpers_in_expr(left, params)),
            right: Box::new(rewrite_php_runtime_arg_helpers_in_expr(right, params)),
        },
        ExprKind::Spread(expr) => ExprKind::Spread(Box::new(
            rewrite_php_runtime_arg_helpers_in_expr(expr, params),
        )),
        ExprKind::Await(expr) => ExprKind::Await(Box::new(
            rewrite_php_runtime_arg_helpers_in_expr(expr, params),
        )),
        ExprKind::Yield(expr) => ExprKind::Yield(
            expr.as_ref()
                .map(|inner| Box::new(rewrite_php_runtime_arg_helpers_in_expr(inner, params))),
        ),
        ExprKind::YieldFrom(expr) => ExprKind::YieldFrom(Box::new(
            rewrite_php_runtime_arg_helpers_in_expr(expr, params),
        )),
        ExprKind::Comprehension {
            kind,
            element,
            generators,
        } => ExprKind::Comprehension {
            kind: *kind,
            element: Box::new(rewrite_php_runtime_arg_helpers_in_expr(element, params)),
            generators: generators.clone(),
        },
        ExprKind::Slice { lower, upper, step } => ExprKind::Slice {
            lower: lower
                .as_ref()
                .map(|inner| Box::new(rewrite_php_runtime_arg_helpers_in_expr(inner, params))),
            upper: upper
                .as_ref()
                .map(|inner| Box::new(rewrite_php_runtime_arg_helpers_in_expr(inner, params))),
            step: step
                .as_ref()
                .map(|inner| Box::new(rewrite_php_runtime_arg_helpers_in_expr(inner, params))),
        },
        ExprKind::Walrus { target, value } => ExprKind::Walrus {
            target: Box::new(rewrite_php_runtime_arg_helpers_in_expr(target, params)),
            value: Box::new(rewrite_php_runtime_arg_helpers_in_expr(value, params)),
        },
        ExprKind::Void(expr) => ExprKind::Void(Box::new(rewrite_php_runtime_arg_helpers_in_expr(
            expr, params,
        ))),
        ExprKind::Delete(expr) => ExprKind::Delete(Box::new(
            rewrite_php_runtime_arg_helpers_in_expr(expr, params),
        )),
        ExprKind::Sequence(exprs) => ExprKind::Sequence(
            exprs
                .iter()
                .map(|inner| rewrite_php_runtime_arg_helpers_in_expr(inner, params))
                .collect(),
        ),
        ExprKind::Range {
            start,
            end,
            inclusive,
        } => ExprKind::Range {
            start: Box::new(rewrite_php_runtime_arg_helpers_in_expr(start, params)),
            end: Box::new(rewrite_php_runtime_arg_helpers_in_expr(end, params)),
            inclusive: *inclusive,
        },
        ExprKind::StaticAccess { class, member } => ExprKind::StaticAccess {
            class: Box::new(rewrite_php_runtime_arg_helpers_in_expr(class, params)),
            member: Box::new(rewrite_php_runtime_arg_helpers_in_expr(member, params)),
        },
        ExprKind::Match { subject, arms } => ExprKind::Match {
            subject: Box::new(rewrite_php_runtime_arg_helpers_in_expr(subject, params)),
            arms: arms
                .iter()
                .map(|arm| MatchArm {
                    conditions: arm.conditions.as_ref().map(|conditions| {
                        conditions
                            .iter()
                            .map(|expr| rewrite_php_runtime_arg_helpers_in_expr(expr, params))
                            .collect()
                    }),
                    body: rewrite_php_runtime_arg_helpers_in_expr(&arm.body, params),
                })
                .collect(),
        },
        _ => expr.kind.clone(),
    };
    Expression::with_span(kind, span)
}

// ─── Literals ─────────────────────────────────────────────────────────────

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().unwrap();
    let kind = match inner.as_rule() {
        Rule::number_lit => walk_number(&inner).kind,
        Rule::string_lit => walk_string(&inner).kind,
        Rule::kw_null => ExprKind::Lit(Literal::Null),
        Rule::kw_true => ExprKind::Lit(Literal::Bool(true)),
        Rule::kw_false => ExprKind::Lit(Literal::Bool(false)),
        _ => ExprKind::Lit(Literal::Null),
    };
    Ok(Expression::with_span(kind, span))
}

fn walk_number(pair: &Pair<Rule>) -> Expression {
    let raw = pair.as_str();
    let s = raw.replace('_', "");
    let kind = if let Some(rest) = s.strip_prefix("0x").or_else(|| s.strip_prefix("0X")) {
        i64::from_str_radix(rest, 16)
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if let Some(rest) = s.strip_prefix("0b").or_else(|| s.strip_prefix("0B")) {
        i64::from_str_radix(rest, 2)
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if let Some(rest) = s.strip_prefix("0o").or_else(|| s.strip_prefix("0O")) {
        i64::from_str_radix(rest, 8)
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if s.len() > 1 && s.starts_with('0') && s.chars().all(|c| ('0'..='7').contains(&c)) {
        i64::from_str_radix(&s[1..], 8)
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if s.contains('.') || s.contains('e') || s.contains('E') {
        s.parse::<f64>()
            .map(Literal::Float)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else {
        s.parse::<i64>()
            .map(Literal::Int)
            .map(ExprKind::Lit)
            .unwrap_or(ExprKind::Lit(Literal::Int(0)))
    };
    Expression::new(kind)
}

fn walk_string(pair: &Pair<Rule>) -> Expression {
    let raw = pair.as_str();

    // Heredoc / nowdoc: `<<<TAG\n...content...\nTAG`
    // Nowdoc uses `<<<'TAG'` — no interpolation.
    if raw.starts_with("<<<") {
        let rest = &raw[3..];
        let is_nowdoc = rest.starts_with('\'');
        // Skip optional quote around tag name
        let tag_start = if is_nowdoc || rest.starts_with('"') {
            &rest[1..]
        } else {
            rest
        };
        let tag_end = tag_start
            .find(|c: char| !c.is_alphanumeric() && c != '_')
            .unwrap_or(tag_start.len());
        let _tag = &tag_start[..tag_end];
        // Content starts after the first newline following the tag line
        let header_end = raw.find('\n').map(|i| i + 1).unwrap_or(raw.len());
        let content_raw = &raw[header_end..];
        // Find the closing tag line. It appears at the start of a line,
        // optionally preceded by whitespace (flexible heredoc PHP 7.3+).
        // Strip the closing tag line — content is everything before the last line
        let content = if let Some(pos) = content_raw.rfind('\n') {
            &content_raw[..pos]
        } else {
            ""
        };
        let content = content.to_string();

        if is_nowdoc {
            return Expression::new(ExprKind::Lit(Literal::Str(unmask_php_literal_tags(
                &content,
            ))));
        }
        // Heredoc: interpolate like double-quoted string
        if !content.contains('$') {
            return Expression::new(ExprKind::Lit(Literal::Str(unmask_php_literal_tags(
                &decode_php_double_quoted_literal(&content),
            ))));
        }
        let parts = parse_php_interpolation(&content);
        if parts.len() == 1 {
            if let InterpolPart::Text(s) = &parts[0] {
                return Expression::new(ExprKind::Lit(Literal::Str(s.clone())));
            }
        }
        if parts.is_empty() {
            return Expression::new(ExprKind::Lit(Literal::Str(String::new())));
        }
        return Expression::new(ExprKind::Interpolation(parts));
    }

    let body = &raw[1..raw.len() - 1];

    if raw.starts_with('\'') {
        // Single-quoted: literal, only \' and \\ escapes. No
        // interpolation in PHP.
        let mut out = String::with_capacity(body.len());
        let mut chars = body.chars().peekable();
        while let Some(c) = chars.next() {
            if c == '\\' {
                if let Some(&next) = chars.peek() {
                    if next == '\'' || next == '\\' {
                        out.push(chars.next().unwrap());
                        continue;
                    }
                }
            }
            out.push(c);
        }
        return Expression::new(ExprKind::Lit(Literal::Str(unmask_php_literal_tags(&out))));
    }

    if !body.contains('$') {
        return Expression::new(ExprKind::Lit(Literal::Str(unmask_php_literal_tags(
            &decode_php_double_quoted_literal(body),
        ))));
    }

    // Double-quoted: PHP interpolation. Scan for `$var`, `$var[key]`,
    // `$var->prop`, `{$expr}` and split the body into InterpolParts.
    // Empty or interp-free strings collapse back to a plain literal so
    // the compiler's string path stays fast.
    let parts = parse_php_interpolation(body);
    if parts.len() == 1 {
        if let InterpolPart::Text(s) = &parts[0] {
            return Expression::new(ExprKind::Lit(Literal::Str(s.clone())));
        }
    }
    if parts.is_empty() {
        return Expression::new(ExprKind::Lit(Literal::Str(String::new())));
    }
    Expression::new(ExprKind::Interpolation(parts))
}

/// Strip common leading indentation from flexible heredoc content (PHP 7.3+).
/// The closing tag's indentation determines how much to strip from every line.
#[allow(dead_code)]
fn strip_heredoc_indentation(content: &str) -> String {
    // Find the minimum indentation (spaces/tabs) across non-empty lines.
    let min_indent = content
        .lines()
        .filter(|l| !l.trim().is_empty())
        .map(|l| l.len() - l.trim_start_matches(|c| c == ' ' || c == '\t').len())
        .min()
        .unwrap_or(0);
    if min_indent == 0 {
        return content.to_string();
    }
    content
        .lines()
        .map(|l| {
            if l.len() >= min_indent {
                &l[min_indent..]
            } else {
                l
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn decode_php_double_quoted_literal(body: &str) -> String {
    let mut out = String::with_capacity(body.len());
    let mut chars = body.chars().peekable();

    while let Some(c) = chars.next() {
        if c != '\\' {
            out.push(c);
            continue;
        }

        let Some(next) = chars.next() else {
            out.push('\\');
            break;
        };

        match next {
            'n' => out.push('\n'),
            'r' => out.push('\r'),
            't' => out.push('\t'),
            'v' => out.push('\u{000B}'),
            'e' => out.push('\u{001B}'),
            'f' => out.push('\u{000C}'),
            '\\' => out.push('\\'),
            '$' => out.push('$'),
            '"' => out.push('"'),
            'x' => {
                let mut hex = String::new();
                for _ in 0..2 {
                    if chars.peek().map(|c| c.is_ascii_hexdigit()).unwrap_or(false) {
                        hex.push(chars.next().unwrap());
                    }
                }
                if let Ok(n) = u32::from_str_radix(&hex, 16) {
                    if let Some(ch) = char::from_u32(n) {
                        out.push(ch);
                    }
                } else {
                    out.push_str("\\x");
                    out.push_str(&hex);
                }
            }
            'u' if chars.peek() == Some(&'{') => {
                chars.next(); // consume {
                let mut hex = String::new();
                while chars.peek().map(|c| *c != '}').unwrap_or(false) {
                    hex.push(chars.next().unwrap());
                }
                chars.next(); // consume }
                if let Ok(n) = u32::from_str_radix(&hex, 16) {
                    if let Some(ch) = char::from_u32(n) {
                        out.push(ch);
                    }
                } else {
                    out.push_str("\\u{");
                    out.push_str(&hex);
                    out.push('}');
                }
            }
            o @ '0'..='7' => {
                let mut oct = String::from(o);
                for _ in 0..2 {
                    if chars
                        .peek()
                        .map(|c| matches!(c, '0'..='7'))
                        .unwrap_or(false)
                    {
                        oct.push(chars.next().unwrap());
                    }
                }
                if let Ok(n) = u32::from_str_radix(&oct, 8) {
                    if let Some(ch) = char::from_u32(n) {
                        out.push(ch);
                    }
                } else {
                    out.push('\\');
                    out.push_str(&oct);
                }
            }
            other => {
                out.push('\\');
                out.push(other);
            }
        }
    }

    out
}

/// Scan a double-quoted PHP string body into `InterpolPart`s, handling:
///   - escape sequences (`\n`, `\t`, `\"`, `\\`, `\$`, …)
///   - `$var`, `$var_with_underscores`
///   - `$arr[key]` — PHP-classic "unquoted key is a string" rule; digit
///     keys become int literals
///   - `$obj->prop`
///   - `{$arbitrary_expr}` — balanced brace matching; inner text parsed
///     by re-entering the PHP expression rule
fn parse_php_interpolation(body: &str) -> Vec<InterpolPart> {
    let mut parts: Vec<InterpolPart> = Vec::new();
    let mut text = String::new();
    let mut chars = body.chars().peekable();

    let flush = |parts: &mut Vec<InterpolPart>, text: &mut String| {
        if !text.is_empty() {
            parts.push(InterpolPart::Text(unmask_php_literal_tags(
                &std::mem::take(text),
            )));
        }
    };

    while let Some(c) = chars.next() {
        // Escapes — must run before $ detection so `\$name` stays literal.
        if c == '\\' {
            match chars.next() {
                Some('n') => text.push('\n'),
                Some('t') => text.push('\t'),
                Some('r') => text.push('\r'),
                Some('"') => text.push('"'),
                Some('\\') => text.push('\\'),
                Some('$') => text.push('$'),
                Some('{') => text.push('{'),
                Some('0') => text.push('\0'),
                Some('x') => {
                    let mut hex = String::new();
                    for _ in 0..2 {
                        if chars.peek().map(|c| c.is_ascii_hexdigit()).unwrap_or(false) {
                            hex.push(chars.next().unwrap());
                        }
                    }
                    if let Ok(n) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(n) {
                            text.push(ch);
                        }
                    } else {
                        text.push_str("\\x");
                        text.push_str(&hex);
                    }
                }
                Some('u') if chars.peek() == Some(&'{') => {
                    chars.next();
                    let mut hex = String::new();
                    while chars.peek().map(|c| *c != '}').unwrap_or(false) {
                        hex.push(chars.next().unwrap());
                    }
                    chars.next();
                    if let Ok(n) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(n) {
                            text.push(ch);
                        }
                    } else {
                        text.push_str("\\u{");
                        text.push_str(&hex);
                        text.push('}');
                    }
                }
                Some(o @ '1'..='7') => {
                    let mut oct = String::from(o);
                    for _ in 0..2 {
                        if chars
                            .peek()
                            .map(|c| matches!(c, '0'..='7'))
                            .unwrap_or(false)
                        {
                            oct.push(chars.next().unwrap());
                        }
                    }
                    if let Ok(n) = u32::from_str_radix(&oct, 8) {
                        if let Some(ch) = char::from_u32(n) {
                            text.push(ch);
                        }
                    } else {
                        text.push('\\');
                        text.push_str(&oct);
                    }
                }
                Some(other) => {
                    text.push('\\');
                    text.push(other);
                }
                None => text.push('\\'),
            }
            continue;
        }

        // `{$...}` complex form — balanced brace scan, re-parse inner.
        if c == '{' && chars.peek() == Some(&'$') {
            flush(&mut parts, &mut text);
            chars.next(); // consume $
            let mut expr_src = String::from("$");
            let mut depth: i32 = 1;
            let mut in_str: Option<char> = None;
            while let Some(&nc) = chars.peek() {
                chars.next();
                if let Some(q) = in_str {
                    expr_src.push(nc);
                    if nc == '\\' {
                        if let Some(&esc) = chars.peek() {
                            expr_src.push(esc);
                            chars.next();
                        }
                        continue;
                    }
                    if nc == q {
                        in_str = None;
                    }
                    continue;
                }
                if nc == '"' || nc == '\'' {
                    in_str = Some(nc);
                    expr_src.push(nc);
                    continue;
                }
                if nc == '{' {
                    depth += 1;
                    expr_src.push(nc);
                    continue;
                }
                if nc == '}' {
                    depth -= 1;
                    if depth == 0 {
                        break;
                    }
                    expr_src.push(nc);
                    continue;
                }
                expr_src.push(nc);
            }
            match parse_interpol_expression(&expr_src) {
                Ok(expr) => parts.push(InterpolPart::Expr(expr)),
                // Fall back to literal so we don't lose user content on
                // parse failure.
                Err(_) => parts.push(InterpolPart::Text(format!("{{{}}}", expr_src))),
            }
            continue;
        }

        // `$identifier` — possibly followed by `[key]` or `->prop`.
        if c == '$' {
            let peek = chars.peek().copied();
            if matches!(peek, Some(c) if c.is_ascii_alphabetic() || c == '_') {
                flush(&mut parts, &mut text);
                let mut name = String::new();
                while let Some(&nc) = chars.peek() {
                    if nc.is_ascii_alphanumeric() || nc == '_' {
                        name.push(nc);
                        chars.next();
                    } else {
                        break;
                    }
                }
                // PHP variables retain the `$` sigil as part of their
                // canonical identifier (see `strip_dollar` — variables
                // and functions live in separate namespaces). The
                // interpolation parser must use the same `$name` form
                // so the generated AST resolves against the global
                // declared by `$name = ...`.
                let mut expr = Expression::new(ExprKind::Ident(format!("${}", name)));

                // `$var[key]` — simple unquoted key; per PHP's rule,
                // identifiers are string literals, digit-runs are ints.
                if chars.peek() == Some(&'[') {
                    chars.next(); // consume [
                    let mut key_text = String::new();
                    while let Some(&nc) = chars.peek() {
                        if nc == ']' {
                            chars.next();
                            break;
                        }
                        key_text.push(nc);
                        chars.next();
                    }
                    let key_trimmed = key_text.trim();
                    let key_expr = if let Ok(n) = key_trimmed.parse::<i64>() {
                        Expression::new(ExprKind::Lit(Literal::Int(n)))
                    } else {
                        // PHP quirk: `$a[$b]` inside string is a variable
                        // if starts with `$`, else unquoted string.
                        if let Some(inner) = key_trimmed.strip_prefix('$') {
                            Expression::new(ExprKind::Ident(inner.to_string()))
                        } else {
                            Expression::new(ExprKind::Lit(Literal::Str(key_trimmed.to_string())))
                        }
                    };
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(key_expr),
                        null_safe: false,
                    });
                } else if chars.peek() == Some(&'-') {
                    // Look ahead for `->`. If absent, `-` is literal.
                    let mut save = chars.clone();
                    save.next();
                    if save.peek() == Some(&'>') {
                        chars.next(); // -
                        chars.next(); // >
                        let mut prop = String::new();
                        while let Some(&nc) = chars.peek() {
                            if nc.is_ascii_alphanumeric() || nc == '_' {
                                prop.push(nc);
                                chars.next();
                            } else {
                                break;
                            }
                        }
                        if !prop.is_empty() {
                            expr = Expression::new(ExprKind::Member {
                                object: Box::new(expr),
                                field: prop,
                                null_safe: false,
                            });
                        }
                    }
                }

                // Coerce the final interpolated value to its string
                // form via `__toString` if it's an object with that
                // magic method (PHP's Stringable contract).
                parts.push(InterpolPart::Expr(php_tostring_coerce(
                    expr,
                    &Span::default(),
                )));
                continue;
            }
            // Lone `$` before non-identifier — literal dollar.
            text.push(c);
            continue;
        }

        text.push(c);
    }

    flush(&mut parts, &mut text);
    parts
}

/// Re-enter the PHP pest grammar on a `{$...}` inner expression.
fn parse_interpol_expression(src: &str) -> Result<Expression, String> {
    use pest::Parser;
    let mut pairs = super::PhpParser::parse(super::Rule::expression, src)
        .map_err(|e| format!("interpolation expr parse failed: {}", e))?;
    let pair = pairs
        .next()
        .ok_or_else(|| "empty interpolation expression".to_string())?;
    walk_expression(pair)
}

// ─── Helpers ──────────────────────────────────────────────────────────────

/// Normalize PHP function-call argument order into the canonical common-AST
/// convention, which matches the JS / Component-Model shape that the
/// compiler emits. PHP builtins whose signature differs from JS need
/// their args rewritten at the walker layer so the downstream compiler
/// sees ONE canonical shape per operation.
///
/// Entries in the match table:
///   ("php_name", &[arg_indices...]) — each arg_indices entry selects
///   which position in the original PHP call the canonical form takes.
///   E.g. `("array_key_exists", &[1, 0])` means the canonical
///   (container, key) order pulls arg 1 first, arg 0 second.
fn canonicalize_php_call_args(callee: &Expression, args: Vec<Argument>) -> Vec<Argument> {
    let name = match &callee.kind {
        ExprKind::Ident(n) => n.as_str(),
        _ => return args,
    };
    let order: &[usize] = match name {
        // PHP: array_key_exists($key, $arr). Canonical: hasOwn($arr, $key).
        "array_key_exists" | "key_exists" => &[1, 0],
        // PHP: in_array($needle, $haystack). Canonical: includes($arr, $needle).
        // Note: `in_array` has an optional 3rd arg (strict); pass through.
        "in_array" => &[1, 0, 2],
        // PHP: implode($glue, $array) / join($glue, $array).
        // Canonical: array.join(delimiter) → host expects (array, delim).
        // PHP also allows the legacy single-arg form `implode($array)` —
        // when only one arg is present we leave it alone so the array
        // surfaces at index 0 and the host fn falls back to "," default.
        "implode" | "join" if args.len() == 2 => &[1, 0],
        // PHP: explode($delim, $string [, $limit]) — opcode `STR_SPLIT`
        // expects `[string, delim]`; swap so the canonical (string,
        // separator) order reaches the VM.
        "explode" if args.len() >= 2 => &[1, 0, 2],
        _ => return args,
    };
    if args.len() < order.iter().filter(|&&i| i < args.len()).count() {
        return args;
    }
    let mut out = Vec::with_capacity(order.len());
    for &i in order {
        if let Some(a) = args.get(i).cloned() {
            out.push(a);
        }
    }
    out
}

fn bind_this_in_lambda_body(body: &LambdaBody, bound_obj_name: &str) -> LambdaBody {
    match body {
        LambdaBody::Expr(expr) => {
            LambdaBody::Expr(Box::new(bind_this_in_expr(expr, bound_obj_name)))
        }
        LambdaBody::Block(stmts) => LambdaBody::Block(
            stmts
                .iter()
                .map(|stmt| bind_this_in_stmt(stmt, bound_obj_name))
                .collect(),
        ),
    }
}

fn bind_this_in_stmt(stmt: &Statement, bound_obj_name: &str) -> Statement {
    let kind = match &stmt.kind {
        StmtKind::Expr(expr) => StmtKind::Expr(bind_this_in_expr(expr, bound_obj_name)),
        StmtKind::Block(body) => StmtKind::Block(
            body.iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
        ),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => StmtKind::If {
            cond: bind_this_in_expr(cond, bound_obj_name),
            then_body: then_body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
            elifs: elifs
                .iter()
                .map(|(cond, body)| {
                    (
                        bind_this_in_expr(cond, bound_obj_name),
                        body.iter()
                            .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                            .collect(),
                    )
                })
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                    .collect()
            }),
        },
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => StmtKind::For {
            init: init
                .as_ref()
                .map(|inner| Box::new(bind_this_in_stmt(inner, bound_obj_name))),
            cond: cond
                .as_ref()
                .map(|expr| bind_this_in_expr(expr, bound_obj_name)),
            update: update
                .as_ref()
                .map(|expr| bind_this_in_expr(expr, bound_obj_name)),
            body: body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
        },
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => StmtKind::ForIn {
            var: var.clone(),
            key: key.clone(),
            iter: bind_this_in_expr(iter, bound_obj_name),
            body: body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
            of: *of,
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                    .collect()
            }),
            is_async: *is_async,
        },
        StmtKind::While {
            cond,
            body,
            else_body,
        } => StmtKind::While {
            cond: bind_this_in_expr(cond, bound_obj_name),
            body: body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                    .collect()
            }),
        },
        StmtKind::DoWhile { body, cond, until } => StmtKind::DoWhile {
            body: body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
            cond: bind_this_in_expr(cond, bound_obj_name),
            until: *until,
        },
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => StmtKind::Switch {
            expr: bind_this_in_expr(expr, bound_obj_name),
            cases: cases
                .iter()
                .map(|case| SwitchCase {
                    conditions: case
                        .conditions
                        .iter()
                        .map(|condition| match condition {
                            CaseCondition::Value(expr) => {
                                CaseCondition::Value(bind_this_in_expr(expr, bound_obj_name))
                            }
                            CaseCondition::Range { from, to } => CaseCondition::Range {
                                from: bind_this_in_expr(from, bound_obj_name),
                                to: bind_this_in_expr(to, bound_obj_name),
                            },
                            CaseCondition::Comparison { op, expr } => CaseCondition::Comparison {
                                op: *op,
                                expr: bind_this_in_expr(expr, bound_obj_name),
                            },
                        })
                        .collect(),
                    body: case
                        .body
                        .iter()
                        .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                        .collect(),
                })
                .collect(),
            default: default.as_ref().map(|body| {
                body.iter()
                    .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                    .collect()
            }),
        },
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => StmtKind::Try {
            body: body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
            catches: catches
                .iter()
                .map(|catch| CatchClause {
                    types: catch.types.clone(),
                    var_name: catch.var_name.clone(),
                    stack_var: catch.stack_var.clone(),
                    body: catch
                        .body
                        .iter()
                        .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                        .collect(),
                    when_clause: catch
                        .when_clause
                        .as_ref()
                        .map(|expr| bind_this_in_expr(expr, bound_obj_name)),
                })
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                    .collect()
            }),
            finally: finally.as_ref().map(|body| {
                body.iter()
                    .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                    .collect()
            }),
        },
        StmtKind::With {
            items,
            body,
            is_async,
        } => StmtKind::With {
            items: items
                .iter()
                .map(|item| WithItem {
                    expr: bind_this_in_expr(&item.expr, bound_obj_name),
                    var: item.var.clone(),
                })
                .collect(),
            body: body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
            is_async: *is_async,
        },
        StmtKind::Using {
            var,
            resource,
            body,
        } => StmtKind::Using {
            var: var.clone(),
            resource: bind_this_in_expr(resource, bound_obj_name),
            body: body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
        },
        StmtKind::Lock { expr, body } => StmtKind::Lock {
            expr: bind_this_in_expr(expr, bound_obj_name),
            body: body
                .iter()
                .map(|inner| bind_this_in_stmt(inner, bound_obj_name))
                .collect(),
        },
        StmtKind::Return(expr) => StmtKind::Return(
            expr.as_ref()
                .map(|inner| bind_this_in_expr(inner, bound_obj_name)),
        ),
        StmtKind::Throw { expr, cause } => StmtKind::Throw {
            expr: expr
                .as_ref()
                .map(|inner| bind_this_in_expr(inner, bound_obj_name)),
            cause: cause
                .as_ref()
                .map(|inner| bind_this_in_expr(inner, bound_obj_name)),
        },
        StmtKind::Assign { targets, value } => StmtKind::Assign {
            targets: targets
                .iter()
                .map(|target| bind_this_in_expr(target, bound_obj_name))
                .collect(),
            value: bind_this_in_expr(value, bound_obj_name),
        },
        StmtKind::CompoundAssign { target, op, value } => StmtKind::CompoundAssign {
            target: bind_this_in_expr(target, bound_obj_name),
            op: *op,
            value: bind_this_in_expr(value, bound_obj_name),
        },
        _ => stmt.kind.clone(),
    };
    Statement::with_span(kind, stmt.span)
}

fn bind_this_in_expr(expr: &Expression, bound_obj_name: &str) -> Expression {
    let span = expr.span;
    let kind = match &expr.kind {
        ExprKind::This => ExprKind::Ident(bound_obj_name.to_string()),
        ExprKind::Ident(name) if name == "$this" => ExprKind::Ident(bound_obj_name.to_string()),
        ExprKind::Binary { op, left, right } => ExprKind::Binary {
            op: *op,
            left: Box::new(bind_this_in_expr(left, bound_obj_name)),
            right: Box::new(bind_this_in_expr(right, bound_obj_name)),
        },
        ExprKind::Unary { op, expr } => ExprKind::Unary {
            op: *op,
            expr: Box::new(bind_this_in_expr(expr, bound_obj_name)),
        },
        ExprKind::Ternary { cond, then, else_ } => ExprKind::Ternary {
            cond: Box::new(bind_this_in_expr(cond, bound_obj_name)),
            then: Box::new(bind_this_in_expr(then, bound_obj_name)),
            else_: Box::new(bind_this_in_expr(else_, bound_obj_name)),
        },
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => ExprKind::Member {
            object: Box::new(bind_this_in_expr(object, bound_obj_name)),
            field: field.clone(),
            null_safe: *null_safe,
        },
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => ExprKind::Index {
            object: Box::new(bind_this_in_expr(object, bound_obj_name)),
            index: Box::new(bind_this_in_expr(index, bound_obj_name)),
            null_safe: *null_safe,
        },
        ExprKind::Call {
            callee,
            args,
            optional,
        } => ExprKind::Call {
            callee: Box::new(bind_this_in_expr(callee, bound_obj_name)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: bind_this_in_expr(&arg.value, bound_obj_name),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        },
        ExprKind::New { class, args } => ExprKind::New {
            class: Box::new(bind_this_in_expr(class, bound_obj_name)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: bind_this_in_expr(&arg.value, bound_obj_name),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
        },
        ExprKind::Assign { target, value } => ExprKind::Assign {
            target: Box::new(bind_this_in_expr(target, bound_obj_name)),
            value: Box::new(bind_this_in_expr(value, bound_obj_name)),
        },
        ExprKind::Array(elements) => ExprKind::Array(
            elements
                .iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .as_ref()
                        .map(|key| bind_this_in_expr(key, bound_obj_name)),
                    value: bind_this_in_expr(&element.value, bound_obj_name),
                    spread: element.spread,
                    by_ref: element.by_ref,
                })
                .collect(),
        ),
        ExprKind::Tuple(items) => ExprKind::Tuple(
            items
                .iter()
                .map(|item| bind_this_in_expr(item, bound_obj_name))
                .collect(),
        ),
        ExprKind::Set(items) => ExprKind::Set(
            items
                .iter()
                .map(|item| bind_this_in_expr(item, bound_obj_name))
                .collect(),
        ),
        ExprKind::Object(props) => ExprKind::Object(
            props
                .iter()
                .map(|prop| match prop {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: bind_this_in_expr(key, bound_obj_name),
                        value: bind_this_in_expr(value, bound_obj_name),
                    },
                    ObjectProperty::Spread(expr) => {
                        ObjectProperty::Spread(bind_this_in_expr(expr, bound_obj_name))
                    }
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: bind_this_in_expr(key, bound_obj_name),
                        value: bind_this_in_expr(value, bound_obj_name),
                    },
                    other => other.clone(),
                })
                .collect(),
        ),
        ExprKind::Interpolation(parts) => ExprKind::Interpolation(
            parts
                .iter()
                .map(|part| match part {
                    InterpolPart::Expr(expr) => {
                        InterpolPart::Expr(bind_this_in_expr(expr, bound_obj_name))
                    }
                    InterpolPart::Formatted(expr, fmt) => InterpolPart::Formatted(
                        bind_this_in_expr(expr, bound_obj_name),
                        fmt.clone(),
                    ),
                    other => other.clone(),
                })
                .collect(),
        ),
        ExprKind::IsType { expr, type_name } => ExprKind::IsType {
            expr: Box::new(bind_this_in_expr(expr, bound_obj_name)),
            type_name: type_name.clone(),
        },
        ExprKind::Cast { expr, type_name } => ExprKind::Cast {
            expr: Box::new(bind_this_in_expr(expr, bound_obj_name)),
            type_name: type_name.clone(),
        },
        ExprKind::TypeOf(expr) => {
            ExprKind::TypeOf(Box::new(bind_this_in_expr(expr, bound_obj_name)))
        }
        ExprKind::NullCoalesce { left, right } => ExprKind::NullCoalesce {
            left: Box::new(bind_this_in_expr(left, bound_obj_name)),
            right: Box::new(bind_this_in_expr(right, bound_obj_name)),
        },
        ExprKind::Spread(expr) => {
            ExprKind::Spread(Box::new(bind_this_in_expr(expr, bound_obj_name)))
        }
        ExprKind::Await(expr) => ExprKind::Await(Box::new(bind_this_in_expr(expr, bound_obj_name))),
        ExprKind::Yield(expr) => ExprKind::Yield(
            expr.as_ref()
                .map(|inner| Box::new(bind_this_in_expr(inner, bound_obj_name))),
        ),
        ExprKind::YieldFrom(expr) => {
            ExprKind::YieldFrom(Box::new(bind_this_in_expr(expr, bound_obj_name)))
        }
        ExprKind::Comprehension {
            kind,
            element,
            generators,
        } => ExprKind::Comprehension {
            kind: *kind,
            element: Box::new(bind_this_in_expr(element, bound_obj_name)),
            generators: generators.clone(),
        },
        ExprKind::Slice { lower, upper, step } => ExprKind::Slice {
            lower: lower
                .as_ref()
                .map(|inner| Box::new(bind_this_in_expr(inner, bound_obj_name))),
            upper: upper
                .as_ref()
                .map(|inner| Box::new(bind_this_in_expr(inner, bound_obj_name))),
            step: step
                .as_ref()
                .map(|inner| Box::new(bind_this_in_expr(inner, bound_obj_name))),
        },
        ExprKind::Walrus { target, value } => ExprKind::Walrus {
            target: Box::new(bind_this_in_expr(target, bound_obj_name)),
            value: Box::new(bind_this_in_expr(value, bound_obj_name)),
        },
        ExprKind::Void(expr) => ExprKind::Void(Box::new(bind_this_in_expr(expr, bound_obj_name))),
        ExprKind::Delete(expr) => {
            ExprKind::Delete(Box::new(bind_this_in_expr(expr, bound_obj_name)))
        }
        ExprKind::Sequence(exprs) => ExprKind::Sequence(
            exprs
                .iter()
                .map(|inner| bind_this_in_expr(inner, bound_obj_name))
                .collect(),
        ),
        ExprKind::Range {
            start,
            end,
            inclusive,
        } => ExprKind::Range {
            start: Box::new(bind_this_in_expr(start, bound_obj_name)),
            end: Box::new(bind_this_in_expr(end, bound_obj_name)),
            inclusive: *inclusive,
        },
        ExprKind::StaticAccess { class, member } => ExprKind::StaticAccess {
            class: Box::new(bind_this_in_expr(class, bound_obj_name)),
            member: Box::new(bind_this_in_expr(member, bound_obj_name)),
        },
        ExprKind::Match { subject, arms } => ExprKind::Match {
            subject: Box::new(bind_this_in_expr(subject, bound_obj_name)),
            arms: arms
                .iter()
                .map(|arm| MatchArm {
                    conditions: arm.conditions.as_ref().map(|conditions| {
                        conditions
                            .iter()
                            .map(|expr| bind_this_in_expr(expr, bound_obj_name))
                            .collect()
                    }),
                    body: bind_this_in_expr(&arm.body, bound_obj_name),
                })
                .collect(),
        },
        _ => expr.kind.clone(),
    };
    Expression::with_span(kind, span)
}

/// Variable identifier passthrough.
///
/// PHP keeps function names and variable names in **separate namespaces**:
/// `function foo() {}` and `$foo = 1` coexist with no collision (call
/// is `foo(...)`, read is `$foo`). Vybex lowers both to global slots
/// keyed by string, so the only way to preserve PHP's two-namespace
/// semantics is to keep the `$` sigil as part of the variable
/// identifier — `$foo` becomes the canonical name, distinct from the
/// bare function name `foo`.
///
/// This started life as `s.strip_prefix('$')` which collapsed both
/// PHP namespaces into one and broke real scripts that legitimately
/// used the same word as both a function and a variable (e.g.
/// `function translate(...)` plus `$translate = Array(...)` in the
/// snif index.php). Returning `s` verbatim is the fix.
fn strip_dollar(s: &str) -> &str {
    s
}

/// Rewrites a PHP function call into the JS-shaped equivalent AST when the
/// callee name maps to a JS standard library function. Returns `None` to
/// leave the call untouched.
///
/// This is the central place where PHP-specific function names get folded
/// into the common AST. Anything that returns Some here flows through the
/// shared compile path with no PHP-aware logic in the compiler / emitter —
/// the same path JS uses (`Math.trunc(...)`, `parseInt(s, base)`, etc.).
fn rewrite_php_call_to_js(callee: &Expression, args: &[Argument], span: &Span) -> Option<ExprKind> {
    let name = match &callee.kind {
        ExprKind::Ident(n) => n.as_str(),
        _ => return None,
    };
    // Helpers for AST construction.
    let mk_lit_f64 = |v: f64| Expression::with_span(ExprKind::Lit(Literal::Float(v)), span.clone());
    let mk_lit_i64 = |v: i64| Expression::with_span(ExprKind::Lit(Literal::Int(v)), span.clone());
    let mk_member = |obj: &str, field: &str| {
        Expression::with_span(
            ExprKind::Member {
                object: Box::new(Expression::with_span(
                    ExprKind::Ident(obj.to_string()),
                    span.clone(),
                )),
                field: field.to_string(),
                null_safe: false,
            },
            span.clone(),
        )
    };
    let mk_binary = |op: BinOp, l: Expression, r: Expression| {
        Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(l),
                right: Box::new(r),
            },
            span.clone(),
        )
    };
    let mk_call = |callee: Expression, call_args: Vec<Expression>| ExprKind::Call {
        callee: Box::new(callee),
        args: call_args.into_iter().map(Argument::positional).collect(),
        optional: false,
    };
    // Extract a positional argument by index — preserves the original
    // expression so spread/by_ref/etc. flags carry through.
    let arg = |i: usize| args.get(i).map(|a| a.value.clone());

    Some(match name {
        // PHP `array_map(null, a, b, ...)` zips arrays into tuples.
        "array_map"
            if args.len() >= 3
                && matches!(
                    args.first().map(|a| &a.value.kind),
                    Some(ExprKind::Lit(Literal::Null))
                ) =>
        {
            mk_call(
                Expression::ident("zip"),
                args.iter().skip(1).map(|arg| arg.value.clone()).collect(),
            )
        }
        // ── Dynamic callable helpers ───────────────────────────────────
        // PHP `call_user_func($cb, ...)` and `call_user_func_array($cb, $args)`
        // are just indirection over the normal call surface. Rewrite them
        // to direct calls so closures, first-class callables, string names,
        // and PHP named-unpack arrays all flow through the existing compiler
        // call logic.
        "call_user_func" if !args.is_empty() => ExprKind::Call {
            callee: Box::new(php_callable_target_expr(arg(0)?, span)),
            args: args.iter().skip(1).cloned().collect(),
            optional: false,
        },
        "call_user_func_array" if args.len() == 2 => ExprKind::Call {
            callee: Box::new(php_callable_target_expr(arg(0)?, span)),
            args: vec![Argument {
                name: None,
                value: arg(1)?,
                by_ref: false,
                spread: true,
            }],
            optional: false,
        },
        // ── Trigonometry / radians ──────────────────────────────────────
        // PHP `deg2rad($x)` → `$x * Math.PI / 180`.
        "deg2rad" => {
            let x = arg(0)?;
            let mul = mk_binary(BinOp::Mul, x, mk_member("Math", "PI"));
            mk_binary(BinOp::Div, mul, mk_lit_f64(180.0)).kind
        }
        // PHP `rad2deg($x)` → `$x * 180 / Math.PI`.
        "rad2deg" => {
            let x = arg(0)?;
            let mul = mk_binary(BinOp::Mul, x, mk_lit_f64(180.0));
            mk_binary(BinOp::Div, mul, mk_member("Math", "PI")).kind
        }
        // ── Integer division ────────────────────────────────────────────
        // PHP `intdiv($a, $b)` truncates toward zero, matching JS
        // `Math.trunc($a / $b)` (NOT Math.floor — different on negatives).
        "intdiv" => {
            let a = arg(0)?;
            let b = arg(1)?;
            let div = mk_binary(BinOp::Div, a, b);
            mk_call(mk_member("Math", "trunc"), vec![div])
        }
        // ── Base conversions: string → integer ──────────────────────────
        // PHP `bindec`/`octdec`/`hexdec` → JS `Number.parseInt(str, base)`.
        // Qualified form (Number.parseInt) bypasses the common-import
        // resolver's bare-`parseInt` → `cint` mapping, which is 1-arg
        // floor-coercion and silently discards the radix.
        "bindec" => mk_call(
            mk_member("Number", "parseInt"),
            vec![arg(0)?, mk_lit_i64(2)],
        ),
        "octdec" => mk_call(
            mk_member("Number", "parseInt"),
            vec![arg(0)?, mk_lit_i64(8)],
        ),
        "hexdec" => mk_call(
            mk_member("Number", "parseInt"),
            vec![arg(0)?, mk_lit_i64(16)],
        ),
        // sprintf is handled entirely by the generic inline formatter
        // (`common:sprintf.format` → emitter::sprintf::build_sprintf). The
        // earlier per-specifier walker rewrites were redundant with — and
        // buggier than — that single path (e.g. the %e/%E str_replace chain
        // doubled the sign on positive exponents), so they were removed in
        // favour of one implementation.
        // ── Type predicates ─────────────────────────────────────────────
        // Map onto JS-shaped `typeof x === "..."` / `Number.isXxx(x)`.
        // PHP-specific receiver classification (array vs object) is
        // deferred — the runtime model needs a separate is-PHP-array
        // tag before we can rewrite `is_array`/`is_object` cleanly.
        "is_string" => {
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        ExprKind::TypeOf(Box::new(arg(0)?)),
                        span.clone(),
                    )),
                    right: Box::new(Expression::with_span(
                        ExprKind::Lit(Literal::Str("string".to_string())),
                        span.clone(),
                    )),
                },
                span.clone(),
            )
            .kind
        }
        "is_bool" => {
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        ExprKind::TypeOf(Box::new(arg(0)?)),
                        span.clone(),
                    )),
                    right: Box::new(Expression::with_span(
                        ExprKind::Lit(Literal::Str("boolean".to_string())),
                        span.clone(),
                    )),
                },
                span.clone(),
            )
            .kind
        }
        "is_callable" => {
            // PHP `is_callable($x)` matches:
            //   - actual functions/closures (typeof === "function")
            //   - string function names that resolve via function_exists()
            //   - objects implementing `__invoke` magic method
            //
            // Walker emits:
            //   typeof $x === "function" ||
            //   (typeof $x === "string" && function_exists($x)) ||
            //   (typeof $x === "object" &&
            //   typeof $x->__invoke === "function")
            //
            // The double-typeof check on the same expression is fine —
            // `arg(0)?` returns a freshly-built clone each call.
            let left_typeof = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        ExprKind::TypeOf(Box::new(arg(0)?)),
                        span.clone(),
                    )),
                    right: Box::new(Expression::string("function")),
                },
                span.clone(),
            );
            let is_string = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        ExprKind::TypeOf(Box::new(arg(0)?)),
                        span.clone(),
                    )),
                    right: Box::new(Expression::string("string")),
                },
                span.clone(),
            );
            let string_callable = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(is_string),
                    right: Box::new(Expression::with_span(
                        ExprKind::Call {
                            callee: Box::new(Expression::ident("function_exists")),
                            args: vec![Argument::positional(arg(0)?)],
                            optional: false,
                        },
                        span.clone(),
                    )),
                },
                span.clone(),
            );
            let is_obj = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        ExprKind::TypeOf(Box::new(arg(0)?)),
                        span.clone(),
                    )),
                    right: Box::new(Expression::string("object")),
                },
                span.clone(),
            );
            let invoke_member = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(arg(0)?),
                    field: "__invoke".to_string(),
                    null_safe: false,
                },
                span.clone(),
            );
            let invoke_typeof = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        ExprKind::TypeOf(Box::new(invoke_member)),
                        span.clone(),
                    )),
                    right: Box::new(Expression::string("function")),
                },
                span.clone(),
            );
            let obj_with_invoke = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(is_obj),
                    right: Box::new(invoke_typeof),
                },
                span.clone(),
            );
            ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::Or,
                        left: Box::new(left_typeof),
                        right: Box::new(string_callable),
                    },
                    span.clone(),
                )),
                right: Box::new(obj_with_invoke),
            }
        }
        "is_null" => {
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(arg(0)?),
                    right: Box::new(Expression::with_span(
                        ExprKind::Lit(Literal::Null),
                        span.clone(),
                    )),
                },
                span.clone(),
            )
            .kind
        }
        "is_int" | "is_integer" | "is_long" => {
            mk_call(mk_member("Number", "isInteger"), vec![arg(0)?])
        }
        "is_finite" => mk_call(mk_member("Number", "isFinite"), vec![arg(0)?]),
        "is_nan" => mk_call(mk_member("Number", "isNaN"), vec![arg(0)?]),
        // PHP `is_infinite($x)` ≡ `Math.abs($x) === Infinity`.
        // `$x` is evaluated once because Math.abs receives it as an
        // argument; the comparison sees only the result.
        // ── Class reflection ────────────────────────────────────────────
        // PHP `get_class($obj)` → `$obj.constructor.name`. Instances carry
        // a `constructor` link to their runtime class (prototype chain in
        // the JS path; stamped directly in the PHP ctor chunk), and the
        // class function carries its declared `name`.
        "get_class" if args.len() == 1 => {
            let ctor = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(arg(0)?),
                    field: "constructor".to_string(),
                    null_safe: false,
                },
                span.clone(),
            );
            ExprKind::Member {
                object: Box::new(ctor),
                field: "name".to_string(),
                null_safe: false,
            }
        }
        // PHP `is_a($obj, "Name")` (literal class name) → `$obj instanceof Name`.
        "is_a" if args.len() == 2 => {
            if let ExprKind::Lit(Literal::Str(class_name)) = &args[1].value.kind {
                ExprKind::Binary {
                    op: BinOp::InstanceOf,
                    left: Box::new(arg(0)?),
                    right: Box::new(Expression::with_span(
                        ExprKind::Ident(class_name.clone()),
                        span.clone(),
                    )),
                }
            } else {
                return None;
            }
        }
        // PHP `is_subclass_of($obj, "Name")` (literal class name) →
        // `$obj instanceof Name && $obj.constructor.name !== "Name"`.
        "is_subclass_of" if args.len() == 2 => {
            if let ExprKind::Lit(Literal::Str(class_name)) = &args[1].value.kind {
                let inst = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::InstanceOf,
                        left: Box::new(arg(0)?),
                        right: Box::new(Expression::with_span(
                            ExprKind::Ident(class_name.clone()),
                            span.clone(),
                        )),
                    },
                    span.clone(),
                );
                let ctor = Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(arg(0)?),
                        field: "constructor".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                );
                let own_name = Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(ctor),
                        field: "name".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                );
                let not_same = mk_binary(
                    BinOp::StrictNotEq,
                    own_name,
                    Expression::string(class_name.as_str()),
                );
                ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(inst),
                    right: Box::new(not_same),
                }
            } else {
                return None;
            }
        }
        // PHP `method_exists($obj, "m")` (literal method name, instance
        // receiver) → `typeof $obj.m === "function"` — instance methods
        // are bound as properties on the instance.
        "method_exists" if args.len() == 2 => {
            if let ExprKind::Lit(Literal::Str(method_name)) = &args[1].value.kind {
                let member = Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(arg(0)?),
                        field: method_name.clone(),
                        null_safe: false,
                    },
                    span.clone(),
                );
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        ExprKind::TypeOf(Box::new(member)),
                        span.clone(),
                    )),
                    right: Box::new(Expression::string("function")),
                }
            } else {
                return None;
            }
        }
        // PHP `property_exists($obj, "p")` → `"p" in $obj` (hasOwn).
        "property_exists" if args.len() == 2 => {
            if let ExprKind::Lit(Literal::Str(_)) = &args[1].value.kind {
                ExprKind::Binary {
                    op: BinOp::In,
                    left: Box::new(arg(1)?),
                    right: Box::new(arg(0)?),
                }
            } else {
                return None;
            }
        }
        // PHP `gettype($v)` → IIFE chain mapping JS typeof onto PHP names.
        "gettype" if args.len() == 1 => {
            let mk_str = |s: &str| {
                Expression::with_span(ExprKind::Lit(Literal::Str(s.to_string())), span.clone())
            };
            let v = Expression::with_span(ExprKind::Ident("v".to_string()), span.clone());
            let typeof_v =
                Expression::with_span(ExprKind::TypeOf(Box::new(v.clone())), span.clone());
            let strict_eq = |left: Expression, right: Expression| {
                Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(left),
                        right: Box::new(right),
                    },
                    span.clone(),
                )
            };
            let ternary = |cond: Expression, then: Expression, else_: Expression| {
                Expression::with_span(
                    ExprKind::Ternary {
                        cond: Box::new(cond),
                        then: Box::new(then),
                        else_: Box::new(else_),
                    },
                    span.clone(),
                )
            };
            let is_int_call = Expression::with_span(
                mk_call(mk_member("Number", "isInteger"), vec![v.clone()]),
                span.clone(),
            );
            // typeof v === "number" ? (Number.isInteger(v) ? "integer" : "double") : ...
            let number_arm = ternary(is_int_call, mk_str("integer"), mk_str("double"));
            let null_check = strict_eq(
                v.clone(),
                Expression::with_span(ExprKind::Lit(Literal::Null), span.clone()),
            );
            // Build chain top-down.
            let chain = ternary(
                null_check,
                mk_str("NULL"),
                ternary(
                    strict_eq(typeof_v.clone(), mk_str("string")),
                    mk_str("string"),
                    ternary(
                        strict_eq(typeof_v.clone(), mk_str("boolean")),
                        mk_str("boolean"),
                        ternary(
                            strict_eq(typeof_v.clone(), mk_str("number")),
                            number_arm,
                            ternary(
                                strict_eq(typeof_v.clone(), mk_str("array")),
                                mk_str("array"),
                                ternary(
                                    strict_eq(typeof_v, mk_str("object")),
                                    mk_str("object"),
                                    mk_str("unknown type"),
                                ),
                            ),
                        ),
                    ),
                ),
            );
            let lambda = Expression::with_span(
                ExprKind::Lambda {
                    params: vec![Param {
                        name: "v".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    }],
                    body: LambdaBody::Expr(Box::new(chain)),
                    is_async: false,
                    captures: vec![],
                },
                span.clone(),
            );
            mk_call(lambda, vec![arg(0)?])
        }
        "is_infinite" => {
            Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::with_span(
                        mk_call(mk_member("Math", "abs"), vec![arg(0)?]),
                        span.clone(),
                    )),
                    right: Box::new(Expression::with_span(
                        ExprKind::Ident("Infinity".to_string()),
                        span.clone(),
                    )),
                },
                span.clone(),
            )
            .kind
        }
        // ── Math constant via function form ─────────────────────────────
        // PHP `pi()` returns the same value as `M_PI` — flatten to the
        // Member-access form so it resolves through namespace_constants.
        "pi" if args.is_empty() => mk_member("Math", "PI").kind,
        // ── String padding (STR_PAD_LEFT/RIGHT only) ─────────────────────
        // PHP `str_pad($s, $n, $pad?, $dir?)` ≡ JS `s.padStart`/`padEnd`.
        // STR_PAD_LEFT  (0) → padStart
        // STR_PAD_RIGHT (1) → padEnd  (default when omitted)
        // STR_PAD_BOTH  (2) → no JS equivalent; walker leaves the call
        //                     in place for the existing polyfill.
        // Dynamic $dir → walker leaves the call in place.
        "str_pad" if args.len() >= 2 => {
            let s = arg(0)?;
            let length = arg(1)?;
            let pad_str = arg(2).unwrap_or_else(|| {
                Expression::with_span(ExprKind::Lit(Literal::Str(" ".to_string())), span.clone())
            });
            // Determine direction. Walker rewrites STR_PAD_* to integer
            // literals in wave 1, so by this point the dir arg is
            // either Lit(Int(...)) or omitted.
            let dir = match args.get(3).map(|a| &a.value.kind) {
                None => Some(1i64), // default STR_PAD_RIGHT
                Some(ExprKind::Lit(Literal::Int(n))) => Some(*n),
                _ => None, // dynamic — fall through to polyfill
            };
            match dir {
                Some(0) => mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(s),
                            field: "padStart".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![length, pad_str],
                ),
                Some(1) => mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(s),
                            field: "padEnd".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![length, pad_str],
                ),
                _ => return None,
            }
        }
        // ── Substring count ─────────────────────────────────────────────
        // PHP `substr_count($h, $n)` ≡ `$h.split($n).length - 1`. The
        // 3rd/4th args (offset, length) are PHP-specific and rare —
        // walker leaves those calls to a polyfill. Empty needle in PHP
        // is a TypeError; JS produces (length-1) chars worth of empty
        // matches. Caller code that hits the empty-needle case is
        // already broken in PHP, so the divergence is acceptable.
        "substr_count" if args.len() == 2 => {
            let h = arg(0)?;
            let n = arg(1)?;
            let split_call = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(h),
                            field: "split".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![n],
                ),
                span.clone(),
            );
            let length_member = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(split_call),
                    field: "length".to_string(),
                    null_safe: false,
                },
                span.clone(),
            );
            ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(length_member),
                right: Box::new(mk_lit_i64(1)),
            }
        }
        // ── URL encoding ────────────────────────────────────────────────
        // PHP `rawurlencode` matches RFC 3986 exactly; same as JS
        // `encodeURIComponent`.
        "rawurlencode" => mk_call(
            Expression::with_span(
                ExprKind::Ident("encodeURIComponent".to_string()),
                span.clone(),
            ),
            vec![arg(0)?],
        ),
        // PHP `urlencode` is RFC 1738 — space → "+" instead of "%20".
        // Compose: `encodeURIComponent($s).split("%20").join("+")`.
        // split/join is used over `replaceAll` because the runtime's
        // replaceAll routes through the regex impl (regex metachars in
        // the search string would misfire); split is literal.
        "urlencode" => {
            let enc = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Ident("encodeURIComponent".to_string()),
                        span.clone(),
                    ),
                    vec![arg(0)?],
                ),
                span.clone(),
            );
            let split_call = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(enc),
                            field: "split".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![Expression::with_span(
                        ExprKind::Lit(Literal::Str("%20".to_string())),
                        span.clone(),
                    )],
                ),
                span.clone(),
            );
            mk_call(
                Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(split_call),
                        field: "join".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                ),
                vec![Expression::with_span(
                    ExprKind::Lit(Literal::Str("+".to_string())),
                    span.clone(),
                )],
            )
        }
        // Lower to PHP-specific helpers so runtime coercion stays in the
        // string adapter and `rawurldecode` preserves literal '+' bytes.
        "urldecode" => mk_call(
            Expression::with_span(ExprKind::Ident("__php_urldecode".to_string()), span.clone()),
            vec![arg(0)?],
        ),
        "rawurldecode" => mk_call(
            Expression::with_span(
                ExprKind::Ident("__php_rawurldecode".to_string()),
                span.clone(),
            ),
            vec![arg(0)?],
        ),
        // PHP `stripos($hay, $needle, $offset?)` has the same false-or-index
        // result shape as `strpos`, but compares case-insensitively.
        // Lower both operands in the walker and reuse the existing
        // `__php_strpos` intrinsic so offset handling and false-on-miss
        // semantics stay centralized.
        "stripos" if args.len() == 2 => {
            let lower_call = |index: usize| -> Option<Expression> {
                Some(Expression::with_span(
                    mk_call(
                        Expression::with_span(
                            ExprKind::Ident("strtolower".to_string()),
                            span.clone(),
                        ),
                        vec![arg(index)?],
                    ),
                    span.clone(),
                ))
            };
            mk_call(
                Expression::with_span(ExprKind::Ident("__php_strpos".to_string()), span.clone()),
                vec![lower_call(0)?, lower_call(1)?],
            )
        }
        "stripos" if args.len() >= 3 => {
            let lower_call = |index: usize| -> Option<Expression> {
                Some(Expression::with_span(
                    mk_call(
                        Expression::with_span(
                            ExprKind::Ident("strtolower".to_string()),
                            span.clone(),
                        ),
                        vec![arg(index)?],
                    ),
                    span.clone(),
                ))
            };
            mk_call(
                Expression::with_span(ExprKind::Ident("__php_strpos".to_string()), span.clone()),
                vec![lower_call(0)?, lower_call(1)?, arg(2)?],
            )
        }
        // ── Precision-aware rounding ────────────────────────────────────
        // PHP `round($n)` ≡ `Math.round($n)`.
        "round" if args.len() == 1 => mk_call(mk_member("Math", "round"), vec![arg(0)?]),
        // PHP `round($n, $p)` ≡ `Math.round($n * Math.pow(10, $p)) / Math.pow(10, $p)`.
        // Both args are evaluated twice — Math.pow on the precision is
        // cheap and side-effect-free, $n is typically a variable. The
        // 3rd PHP arg (mode flag) is ignored — round-half-to-even is
        // identical to Math.round on positives, which covers the test
        // suite's needs; banker's rounding can be added later if it
        // turns out to matter.
        "round" if args.len() >= 2 => {
            let n_first = arg(0)?;
            let n_second = arg(0)?;
            let p_first = arg(1)?;
            let p_second = arg(1)?;
            // Math.pow(10, p)
            let pow_first = Expression::with_span(
                mk_call(mk_member("Math", "pow"), vec![mk_lit_f64(10.0), p_first]),
                span.clone(),
            );
            let pow_second = Expression::with_span(
                mk_call(mk_member("Math", "pow"), vec![mk_lit_f64(10.0), p_second]),
                span.clone(),
            );
            // n * pow(10, p)
            let scaled = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::Mul,
                    left: Box::new(n_first),
                    right: Box::new(pow_first),
                },
                span.clone(),
            );
            // Wait — should be n_second... clarify: we evaluated arg twice; first scaled,
            // second is the divisor below. Use n_first for the multiplication.
            let _ = n_second;
            // Math.round(scaled)
            let rounded = Expression::with_span(
                mk_call(mk_member("Math", "round"), vec![scaled]),
                span.clone(),
            );
            // rounded / pow(10, p)
            ExprKind::Binary {
                op: BinOp::Div,
                left: Box::new(rounded),
                right: Box::new(pow_second),
            }
        }
        // ── Letter-case-first ───────────────────────────────────────────
        // PHP `ucfirst($s)` ≡ `$s.charAt(0).toUpperCase() + $s.slice(1)`.
        // `$s` is evaluated twice — typical PHP usage passes a variable.
        "ucfirst" => {
            let s_left = arg(0)?;
            let s_right = arg(0)?;
            let first_char = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(s_left),
                            field: "charAt".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![mk_lit_i64(0)],
                ),
                span.clone(),
            );
            let upper_first = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(first_char),
                            field: "toUpperCase".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![],
                ),
                span.clone(),
            );
            let rest = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(s_right),
                            field: "slice".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![mk_lit_i64(1)],
                ),
                span.clone(),
            );
            ExprKind::Binary {
                op: BinOp::Concat,
                left: Box::new(upper_first),
                right: Box::new(rest),
            }
        }
        // PHP `strcasecmp($a, $b)` is the case-insensitive sibling of
        // `strcmp`. Lower both operands and route through the existing
        // string-compare opcode so direct calls and callback contexts
        // share one implementation surface.
        "strcasecmp" if args.len() == 2 => {
            let lower_arg = |index: usize| -> Option<Expression> {
                Some(Expression::with_span(
                    mk_call(
                        Expression::with_span(
                            ExprKind::Ident("strtolower".to_string()),
                            span.clone(),
                        ),
                        vec![arg(index)?],
                    ),
                    span.clone(),
                ))
            };
            mk_call(
                Expression::with_span(ExprKind::Ident("strcmp".to_string()), span.clone()),
                vec![lower_arg(0)?, lower_arg(1)?],
            )
        }
        // strncmp/strncasecmp handled via profile + emitter (not walker rewrite)
        // fnmatch($pattern, $string) — rewrite to intrinsic
        "fnmatch" if args.len() == 2 => mk_call(
            Expression::with_span(ExprKind::Ident("__vybe_fnmatch".to_string()), span.clone()),
            vec![arg(0)?, arg(1)?],
        ),
        // strtok($str, $delim) / strtok($delim) — stateful tokenizer
        "strtok" if args.len() == 2 => mk_call(
            Expression::with_span(ExprKind::Ident("__vybe_strtok".to_string()), span.clone()),
            vec![arg(0)?, arg(1)?],
        ),
        "strtok" if args.len() == 1 => mk_call(
            Expression::with_span(
                ExprKind::Ident("__vybe_strtok_next".to_string()),
                span.clone(),
            ),
            vec![arg(0)?],
        ),
        // mb_strrpos → strrpos
        "mb_strrpos" if args.len() == 2 => mk_call(
            Expression::with_span(ExprKind::Ident("strrpos".to_string()), span.clone()),
            vec![arg(0)?, arg(1)?],
        ),
        // mb_detect_encoding → always "UTF-8"
        "mb_detect_encoding" => ExprKind::Lit(Literal::Str("UTF-8".to_string())),
        // mb_check_encoding → always true
        "mb_check_encoding" => ExprKind::Lit(Literal::Bool(true)),
        // mb_str_split($s, $n=1) → str_split($s, $n)
        "mb_str_split" if !args.is_empty() => {
            let n = if args.len() >= 2 {
                arg(1)?
            } else {
                mk_lit_i64(1)
            };
            mk_call(
                Expression::with_span(ExprKind::Ident("str_split".to_string()), span.clone()),
                vec![arg(0)?, n],
            )
        }
        // mb_substr_count → substr_count
        "mb_substr_count" if args.len() >= 2 => mk_call(
            Expression::with_span(ExprKind::Ident("substr_count".to_string()), span.clone()),
            vec![arg(0)?, arg(1)?],
        ),
        // mb_convert_case($s, $mode)
        "mb_convert_case" if args.len() >= 2 => mk_call(
            Expression::with_span(
                ExprKind::Ident("__vybe_mb_convert_case".to_string()),
                span.clone(),
            ),
            vec![arg(0)?, arg(1)?],
        ),
        // mb_str_pad → str_pad
        "mb_str_pad" if !args.is_empty() => {
            let rargs: Vec<Expression> = (0..args.len().min(4)).filter_map(|i| arg(i)).collect();
            mk_call(
                Expression::with_span(ExprKind::Ident("str_pad".to_string()), span.clone()),
                rargs,
            )
        }
        // ── PHP higher-order array fns with PHP-specific semantics ────
        // `array_map` preserves associative keys for the 1-array form
        // and `array_filter` has PHP-only flag modes like
        // ARRAY_FILTER_USE_KEY, so both stay on the PHP adapter layer.
        "array_reduce" if args.len() >= 2 => {
            let arr_expr = arg(0)?;
            let cb = arg(1)?;
            let mut call_args = vec![cb];
            if let Some(init) = arg(2) {
                call_args.push(init);
            }
            mk_call(
                Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(arr_expr),
                        field: "reduce".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                ),
                call_args,
            )
        }
        // ── Bare math functions → Math.* member calls ──────────────────
        // PHP exposes these as bare globals; JS spec puts them on Math.
        // Walker rewrites so the compile path goes through the
        // namespace-method dispatch (single source of truth — PHP profile
        // binds `Math.pow`/`Math.sin`/etc. once, no parallel bare entries).
        "abs" if args.len() == 1 => mk_call(mk_member("Math", "abs"), vec![arg(0)?]),
        "sqrt" if args.len() == 1 => mk_call(mk_member("Math", "sqrt"), vec![arg(0)?]),
        "floor" if args.len() == 1 => mk_call(mk_member("Math", "floor"), vec![arg(0)?]),
        "ceil" if args.len() == 1 => mk_call(mk_member("Math", "ceil"), vec![arg(0)?]),
        "pow" if args.len() == 2 => mk_call(mk_member("Math", "pow"), vec![arg(0)?, arg(1)?]),
        "exp" if args.len() == 1 => mk_call(mk_member("Math", "exp"), vec![arg(0)?]),
        "log" if args.len() == 1 => mk_call(mk_member("Math", "log"), vec![arg(0)?]),
        "log2" if args.len() == 1 => mk_call(mk_member("Math", "log2"), vec![arg(0)?]),
        "log10" if args.len() == 1 => mk_call(mk_member("Math", "log10"), vec![arg(0)?]),
        "sin" if args.len() == 1 => mk_call(mk_member("Math", "sin"), vec![arg(0)?]),
        "cos" if args.len() == 1 => mk_call(mk_member("Math", "cos"), vec![arg(0)?]),
        "tan" if args.len() == 1 => mk_call(mk_member("Math", "tan"), vec![arg(0)?]),
        "asin" if args.len() == 1 => mk_call(mk_member("Math", "asin"), vec![arg(0)?]),
        "acos" if args.len() == 1 => mk_call(mk_member("Math", "acos"), vec![arg(0)?]),
        "atan" if args.len() == 1 => mk_call(mk_member("Math", "atan"), vec![arg(0)?]),
        "atan2" if args.len() == 2 => mk_call(mk_member("Math", "atan2"), vec![arg(0)?, arg(1)?]),
        "sinh" if args.len() == 1 => mk_call(mk_member("Math", "sinh"), vec![arg(0)?]),
        "cosh" if args.len() == 1 => mk_call(mk_member("Math", "cosh"), vec![arg(0)?]),
        "tanh" if args.len() == 1 => mk_call(mk_member("Math", "tanh"), vec![arg(0)?]),
        "asinh" if args.len() == 1 => mk_call(mk_member("Math", "asinh"), vec![arg(0)?]),
        "acosh" if args.len() == 1 => mk_call(mk_member("Math", "acosh"), vec![arg(0)?]),
        "atanh" if args.len() == 1 => mk_call(mk_member("Math", "atanh"), vec![arg(0)?]),
        // ── Array splice ───────────────────────────────────────────────
        // PHP `array_splice($arr, $start, $length?, $replacement?)` ≡
        // JS `$arr.splice($start, $length, ...$replacement)`. PHP packs
        // the replacement items into a single array; JS spreads them.
        // Walker emits a Spread argument so the existing JS-shape splice
        // dispatch (no Vybe-specific code needed) handles it.
        "array_splice" if args.len() >= 2 => {
            let arr_expr = arg(0)?;
            let start = arg(1)?;
            // Default $length = remove to end (i32::MAX clamped by JS impl).
            let length = arg(2).unwrap_or_else(|| {
                Expression::with_span(ExprKind::Lit(Literal::Int(i32::MAX as i64)), span.clone())
            });
            let mut splice_args: Vec<Argument> =
                vec![Argument::positional(start), Argument::positional(length)];
            // 4th arg: replacement array. Spread its elements.
            if let Some(rep) = arg(3) {
                splice_args.push(Argument {
                    value: rep,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
            }
            ExprKind::Call {
                callee: Box::new(Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(arr_expr),
                        field: "splice".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                )),
                args: splice_args,
                optional: false,
            }
        }
        // ── Array sum ──────────────────────────────────────────────────
        // PHP `array_sum($arr)` → `$arr.reduce((a, b) => a + b, 0)`.
        "array_sum" => {
            let arr_expr = arg(0)?;
            let body = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(Expression::with_span(
                        ExprKind::Ident("__a".to_string()),
                        span.clone(),
                    )),
                    right: Box::new(Expression::with_span(
                        ExprKind::Ident("__b".to_string()),
                        span.clone(),
                    )),
                },
                span.clone(),
            );
            let lambda = Expression::with_span(
                ExprKind::Lambda {
                    params: vec![
                        Param {
                            name: "__a".to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                        Param {
                            name: "__b".to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                    ],
                    body: LambdaBody::Expr(Box::new(body)),
                    is_async: false,
                    captures: vec![],
                },
                span.clone(),
            );
            mk_call(
                Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(arr_expr),
                        field: "reduce".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                ),
                vec![lambda, mk_lit_f64(0.0)],
            )
        }
        // ── Array product ──────────────────────────────────────────────
        // PHP `array_product($arr)` → `$arr.reduce((a, b) => a * b, 1)`.
        // Resolves through PHP profile's [array_methods] (`reduce` →
        // __array_reduce loop helper). Single-eval — `$arr` appears
        // once; the lambda doesn't capture.
        "array_product" => {
            let arr_expr = arg(0)?;
            let body = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::Mul,
                    left: Box::new(Expression::with_span(
                        ExprKind::Ident("__a".to_string()),
                        span.clone(),
                    )),
                    right: Box::new(Expression::with_span(
                        ExprKind::Ident("__b".to_string()),
                        span.clone(),
                    )),
                },
                span.clone(),
            );
            let lambda = Expression::with_span(
                ExprKind::Lambda {
                    params: vec![
                        Param {
                            name: "__a".to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                        Param {
                            name: "__b".to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                    ],
                    body: LambdaBody::Expr(Box::new(body)),
                    is_async: false,
                    captures: vec![],
                },
                span.clone(),
            );
            mk_call(
                Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(arr_expr),
                        field: "reduce".to_string(),
                        null_safe: false,
                    },
                    span.clone(),
                ),
                vec![lambda, mk_lit_f64(1.0)],
            )
        }
        // PHP `substr($s, $start, $length?)` →
        //   2-arg: `__php_substr($s, $start)`
        //   3-arg: `__php_substr($s, $start, $length)`
        // The compiler intrinsic lowers this directly to STR_SUBSTRING,
        // which keeps dynamic receivers safe (e.g. `$_SERVER[...]`).
        "substr" | "mb_substr" if args.len() == 2 => mk_call(
            Expression::with_span(ExprKind::Ident("__php_substr".to_string()), span.clone()),
            vec![arg(0)?, arg(1)?],
        ),
        "substr" | "mb_substr" if args.len() >= 3 => mk_call(
            Expression::with_span(ExprKind::Ident("__php_substr".to_string()), span.clone()),
            vec![arg(0)?, arg(1)?, arg(2)?],
        ),
        // PHP `lcfirst($s)` — same shape, lowercase first character.
        "lcfirst" => {
            let s_left = arg(0)?;
            let s_right = arg(0)?;
            let first_char = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(s_left),
                            field: "charAt".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![mk_lit_i64(0)],
                ),
                span.clone(),
            );
            let lower_first = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(first_char),
                            field: "toLowerCase".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![],
                ),
                span.clone(),
            );
            let rest = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Member {
                            object: Box::new(s_right),
                            field: "slice".to_string(),
                            null_safe: false,
                        },
                        span.clone(),
                    ),
                    vec![mk_lit_i64(1)],
                ),
                span.clone(),
            );
            ExprKind::Binary {
                op: BinOp::Concat,
                left: Box::new(lower_first),
                right: Box::new(rest),
            }
        }
        // PHP `strtotime("+N unit", $base)` with a literal relative string
        // is rewritten at compile time to `$base + N * unit_secs` (or
        // calendar arithmetic for month/year via `ecma:date.UTC` round-
        // trip). When the string isn't recognised the rewrite returns
        // `None` so the runtime adapter handles it.
        "strtotime" if args.len() == 2 => {
            let s = match &args[0].value.kind {
                ExprKind::Lit(Literal::Str(s)) => s,
                _ => return None,
            };
            let (n, unit) = crate::emitter::php::datetime_adapter::parse_relative_delta(s)?;
            let base = arg(1)?;
            let secs_per_unit: Option<i64> = match unit {
                "second" => Some(1),
                "minute" => Some(60),
                "hour" => Some(3_600),
                "day" => Some(86_400),
                "week" => Some(604_800),
                _ => None, // month / year need calendar arithmetic
            };
            if let Some(secs) = secs_per_unit {
                // base + n * secs_per_unit
                let delta = mk_lit_i64(n * secs);
                mk_binary(BinOp::Add, base, delta).kind
            } else {
                // month/year — calendar arithmetic via the bytecode
                // adapter (`__php_strtotime_rel_calendar(base, n, is_year)`).
                let is_year = unit == "year";
                let bool_arg =
                    Expression::with_span(ExprKind::Lit(Literal::Bool(is_year)), span.clone());
                mk_call(
                    Expression::with_span(
                        ExprKind::Ident("__php_strtotime_rel_calendar".to_string()),
                        span.clone(),
                    ),
                    vec![base, mk_lit_i64(n), bool_arg],
                )
            }
        }
        // PHP `preg_match($pat, $str, $matches)` and `preg_match_all` —
        // PHP populates `$matches` by reference with capture groups
        // and returns the count. Rewrite the 3-arg form to a Sequence
        // expression: `($matches = __preg_match_all_groups(pat, str),
        // count_of_matches)` so both assignment side-effect AND
        // count-as-result behave like real PHP.
        // preg_replace with limit — if limit != -1, use single-replace
        "preg_replace" if args.len() == 4 => {
            // (pat, repl, str, limit) — if limit is -1 use replaceAll, else use replace (first match)
            // For limit != -1, route to __vybe_preg_replace_limited helper
            mk_call(
                Expression::with_span(
                    ExprKind::Ident("__vybe_preg_replace_limited".to_string()),
                    span.clone(),
                ),
                vec![arg(0)?, arg(1)?, arg(2)?, arg(3)?],
            )
        }
        // 2-arg: just return the match count
        "preg_match_all" if args.len() == 2 => {
            let groups_call = Expression::with_span(
                mk_call(
                    Expression::with_span(
                        ExprKind::Ident("__preg_match_all_groups".to_string()),
                        span.clone(),
                    ),
                    vec![arg(0)?, arg(1)?],
                ),
                span.clone(),
            );
            let zero = Expression::with_span(ExprKind::Lit(Literal::Int(0)), span.clone());
            let col0 = Expression::with_span(
                ExprKind::Index {
                    object: Box::new(groups_call),
                    index: Box::new(zero),
                    null_safe: false,
                },
                span.clone(),
            );
            ExprKind::Member {
                object: Box::new(col0),
                field: "length".to_string(),
                null_safe: false,
            }
        }
        "preg_match_all" if args.len() == 3 => {
            let target = args[2].value.clone();
            let groups_call = mk_call(
                Expression::with_span(
                    ExprKind::Ident("__preg_match_all_groups".to_string()),
                    span.clone(),
                ),
                vec![arg(0)?, arg(1)?],
            );
            let assign = Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(target.clone()),
                    value: Box::new(Expression::with_span(groups_call, span.clone())),
                },
                span.clone(),
            );
            // Count = $matches[0].length (length of full-match column).
            let zero = Expression::with_span(ExprKind::Lit(Literal::Int(0)), span.clone());
            let column0 = Expression::with_span(
                ExprKind::Index {
                    object: Box::new(target),
                    index: Box::new(zero),
                    null_safe: false,
                },
                span.clone(),
            );
            let count = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(column0),
                    field: "length".to_string(),
                    null_safe: false,
                },
                span.clone(),
            );
            ExprKind::Sequence(vec![assign, count])
        }
        // 3-arg `preg_match` returns 0 or 1 (match found?). Walker
        // rewrite: assign matches map, then yield 0/1 based on whether
        // the map is non-empty.
        "preg_match" if args.len() == 3 => {
            let target = args[2].value.clone();
            let groups_call = mk_call(
                Expression::with_span(
                    ExprKind::Ident("__preg_match_groups".to_string()),
                    span.clone(),
                ),
                vec![arg(0)?, arg(1)?],
            );
            let assign = Expression::with_span(
                ExprKind::Assign {
                    target: Box::new(target.clone()),
                    value: Box::new(Expression::with_span(groups_call, span.clone())),
                },
                span.clone(),
            );
            // 1 if $matches has at least one numeric entry; else 0.
            // Simplest probe: `$matches[0] !== undefined` ternary → 1/0.
            let zero = Expression::with_span(ExprKind::Lit(Literal::Int(0)), span.clone());
            let m0 = Expression::with_span(
                ExprKind::Index {
                    object: Box::new(target),
                    index: Box::new(zero),
                    null_safe: false,
                },
                span.clone(),
            );
            let undef = Expression::with_span(ExprKind::Lit(Literal::Undefined), span.clone());
            let cmp = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictNotEq,
                    left: Box::new(m0),
                    right: Box::new(undef),
                },
                span.clone(),
            );
            let one = Expression::with_span(ExprKind::Lit(Literal::Int(1)), span.clone());
            let zero2 = Expression::with_span(ExprKind::Lit(Literal::Int(0)), span.clone());
            let count = Expression::with_span(
                ExprKind::Ternary {
                    cond: Box::new(cmp),
                    then: Box::new(one),
                    else_: Box::new(zero2),
                },
                span.clone(),
            );
            ExprKind::Sequence(vec![assign, count])
        }
        _ => return None,
    })
}

/// Rewrites a PHP global constant name (`M_PI`, `STR_PAD_LEFT`, etc.) into
/// the JS-shaped common-AST expression. Returns `None` for non-constants.
///
/// The compiler's `lookup_constant` resolves the resulting `Math.PI` /
/// `Number.MAX_SAFE_INTEGER` Member-access AST against the profile's
/// `[namespace_constants]` table — same path JS uses. No PHP-specific
/// intrinsic, host fn, or profile builtin needed.
fn php_constant_expr(name: &str, span: &Span) -> Option<ExprKind> {
    let mk_member = |obj: &str, field: &str| ExprKind::Member {
        object: Box::new(Expression::with_span(
            ExprKind::Ident(obj.to_string()),
            span.clone(),
        )),
        field: field.to_string(),
        null_safe: false,
    };
    let mk_div = |num: f64, den_obj: &str, den_field: &str| ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(Expression::with_span(
            ExprKind::Lit(Literal::Float(num)),
            span.clone(),
        )),
        right: Box::new(Expression::with_span(
            mk_member(den_obj, den_field),
            span.clone(),
        )),
    };
    let mk_div_by_lit = |num_obj: &str, num_field: &str, den: f64| ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(Expression::with_span(
            mk_member(num_obj, num_field),
            span.clone(),
        )),
        right: Box::new(Expression::with_span(
            ExprKind::Lit(Literal::Float(den)),
            span.clone(),
        )),
    };
    Some(match name {
        // ── math constants → Math.* property ──
        "M_PI" => mk_member("Math", "PI"),
        "M_E" => mk_member("Math", "E"),
        "M_LN2" => mk_member("Math", "LN2"),
        "M_LN10" => mk_member("Math", "LN10"),
        "M_LOG2E" => mk_member("Math", "LOG2E"),
        "M_LOG10E" => mk_member("Math", "LOG10E"),
        "M_SQRT2" => mk_member("Math", "SQRT2"),
        "M_SQRT1_2" => mk_member("Math", "SQRT1_2"),
        // ── math composites — no JS data property, compose ──
        "M_PI_2" => mk_div_by_lit("Math", "PI", 2.0),
        "M_PI_4" => mk_div_by_lit("Math", "PI", 4.0),
        "M_1_PI" => mk_div(1.0, "Math", "PI"),
        "M_2_PI" => mk_div(2.0, "Math", "PI"),
        // M_2_SQRTPI = 2/sqrt(PI) — no clean JS form, bake the literal.
        "M_2_SQRTPI" => ExprKind::Lit(Literal::Float(std::f64::consts::FRAC_2_SQRT_PI)),
        // ── infinity / NaN — JS globals ──
        "INF" => ExprKind::Ident("Infinity".to_string()),
        "NAN" => ExprKind::Ident("NaN".to_string()),
        // ── PHP integer / float limits → Number.* property ──
        "PHP_INT_MAX" => mk_member("Number", "MAX_SAFE_INTEGER"),
        "PHP_INT_MIN" => mk_member("Number", "MIN_SAFE_INTEGER"),
        "PHP_FLOAT_MAX" => mk_member("Number", "MAX_VALUE"),
        "PHP_FLOAT_MIN" => mk_member("Number", "MIN_VALUE"),
        "PHP_FLOAT_EPSILON" => mk_member("Number", "EPSILON"),
        // ── PHP integer-like literals ──
        "PHP_INT_SIZE" => ExprKind::Lit(Literal::Int(8)),
        "PHP_FLOAT_DIG" => ExprKind::Lit(Literal::Int(15)),
        // ── string padding flags — integer literals ──
        "STR_PAD_LEFT" => ExprKind::Lit(Literal::Int(0)),
        "STR_PAD_RIGHT" => ExprKind::Lit(Literal::Int(1)),
        "STR_PAD_BOTH" => ExprKind::Lit(Literal::Int(2)),
        // ── sort flags — integer literals ──
        "SORT_REGULAR" => ExprKind::Lit(Literal::Int(0)),
        "SORT_NUMERIC" => ExprKind::Lit(Literal::Int(1)),
        "SORT_STRING" => ExprKind::Lit(Literal::Int(2)),
        "SORT_DESC" => ExprKind::Lit(Literal::Int(3)),
        "SORT_ASC" => ExprKind::Lit(Literal::Int(4)),
        "SORT_FLAG_CASE" => ExprKind::Lit(Literal::Int(8)),
        "SORT_NATURAL" => ExprKind::Lit(Literal::Int(6)),
        "SORT_LOCALE_STRING" => ExprKind::Lit(Literal::Int(5)),
        // ── filter flags — integer literals ──
        "ARRAY_FILTER_USE_KEY" => ExprKind::Lit(Literal::Int(2)),
        "ARRAY_FILTER_USE_BOTH" => ExprKind::Lit(Literal::Int(1)),
        // ── preg flags ──
        "PREG_GREP_INVERT" => ExprKind::Lit(Literal::Int(1)),
        "PREG_SPLIT_NO_EMPTY" => ExprKind::Lit(Literal::Int(1)),
        "PREG_SPLIT_DELIM_CAPTURE" => ExprKind::Lit(Literal::Int(2)),
        "PREG_SET_ORDER" => ExprKind::Lit(Literal::Int(2)),
        "PREG_OFFSET_CAPTURE" => ExprKind::Lit(Literal::Int(256)),
        // ── mb_ case constants ──
        "MB_CASE_UPPER" => ExprKind::Lit(Literal::Int(0)),
        "MB_CASE_LOWER" => ExprKind::Lit(Literal::Int(1)),
        "MB_CASE_TITLE" => ExprKind::Lit(Literal::Int(2)),
        "MB_CASE_FOLD" => ExprKind::Lit(Literal::Int(80)),
        "MB_CASE_UPPER_SIMPLE" => ExprKind::Lit(Literal::Int(40)),
        "MB_CASE_LOWER_SIMPLE" => ExprKind::Lit(Literal::Int(41)),
        "MB_CASE_FOLD_SIMPLE" => ExprKind::Lit(Literal::Int(81)),
        // ── newline / paths — string literals ──
        "PHP_EOL" => ExprKind::Lit(Literal::Str("\n".to_string())),
        "DIRECTORY_SEPARATOR" => ExprKind::Lit(Literal::Str("/".to_string())),
        "PATH_SEPARATOR" => ExprKind::Lit(Literal::Str(":".to_string())),
        _ => return None,
    })
}

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    LINE_STARTS.with(|starts| {
        let starts = starts.borrow();
        if starts.is_empty() {
            let (start_line, start_col) = s.start_pos().line_col();
            let (end_line, end_col) = s.end_pos().line_col();
            Span {
                start_line: start_line as u32,
                start_col: start_col as u32,
                end_line: end_line as u32,
                end_col: end_col as u32,
            }
        } else {
            let (start_line, start_col) = offset_to_line_col(s.start(), &starts);
            let (end_line, end_col) = offset_to_line_col(s.end(), &starts);
            Span {
                start_line,
                start_col,
                end_line,
                end_col,
            }
        }
    })
}
