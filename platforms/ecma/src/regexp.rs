//! `ecma:regexp` — ECMA-262 §22.2 RegExp + the regex-taking
//! `String.prototype` methods (match, matchAll, search, replace,
//! replaceAll, split).
//!
//! Backed by the `regress` crate, which targets ECMAScript regexp
//! syntax more closely than Rust's common regex engine.
//!
//! JS flag handling (ECMA-262 §22.2.5.1):
//!   `i` / `m` / `s` / `u` / `v` → passed through to `regress`
//!   `g` / `y` → handled by the wrapper (`find_iter` / `lastIndex`)
//!   `d` → not yet surfaced as indices objects; ignored by the wrapper
//!
//! Construct shape:
//!   - `ObjectKind::Ordinary` with properties `source`, `flags`, `global`,
//!     `ignoreCase`, `multiline`, `dotAll`, `unicode`, `sticky`,
//!     `lastIndex`, `__type=RegExp`. The `__type` stamp lets
//!     `instanceof RegExp` work via the cross-language type registry.

use icu::properties::CodePointSetData;
use icu::properties::props::{
    Emoji, EmojiComponent, EmojiModifier, EmojiModifierBase, EmojiPresentation, RegionalIndicator,
};
use regress::{Match, Regex};
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::VecDeque;
use std::fmt::Write as _;
use std::sync::{Arc, Mutex};
use unicode_segmentation::UnicodeSegmentation;
use vybe_runtime::value::{Object, Value};
use vybe_runtime::{HostContext, VM};

const REGEXP_TYPE: &str = "RegExp";
const REGEXP_CACHE_LIMIT: usize = 64;
const REGEXP_REPLACE_INLINE_ARG_LIMIT: usize = 8;

static REGEXP_PROTOTYPE: std::sync::OnceLock<Arc<Mutex<Object>>> = std::sync::OnceLock::new();

thread_local! {
    static REGEXP_CACHE: RefCell<VecDeque<(Arc<str>, Arc<str>, Arc<Regex>)>> =
        RefCell::new(VecDeque::with_capacity(REGEXP_CACHE_LIMIT));
}

#[inline]
fn char_value(ch: char) -> Value {
    crate::keys::char_value(ch)
}

/// %RegExp.prototype% — process-global singleton (same pattern as the
/// other builtin prototypes). Instances link to it via `__proto__`, so
/// `Object.getPrototypeOf(/a/) === RegExp.prototype` and
/// `RegExp.prototype.isPrototypeOf(/a/)` hold (§22.2.6).
pub fn shared_regexp_prototype() -> Value {
    Value::Object(
        REGEXP_PROTOTYPE
            .get_or_init(|| {
                let mut proto = Object::new();
                proto
                    .properties
                    .insert("__proto__".into(), crate::object::shared_object_prototype());
                vybe_runtime::heap::alloc(proto)
            })
            .clone(),
    )
}

#[derive(Clone, Copy)]
enum SpecialPattern {
    RgiEmoji,
}

fn with_s_arg<R>(args: &[Value], idx: usize, f: impl FnOnce(&str) -> R) -> R {
    match args.get(idx) {
        Some(Value::String(text)) => f(text.as_ref()),
        Some(other) => {
            let text = crate::keys::value_display_string(other);
            f(&text)
        }
        None => f(""),
    }
}

fn with_two_s_args<R>(
    args: &[Value],
    first: usize,
    second: usize,
    f: impl FnOnce(&str, &str) -> R,
) -> R {
    with_s_arg(args, first, |a| with_s_arg(args, second, |b| f(a, b)))
}

/// Extract `source` + `flags` from a RegExp object (or treat raw string
/// args as `(pattern, flags)`). Returns `(pattern, flags)` strings.
fn extract_pattern(args: &[Value], idx: usize) -> (String, String) {
    match args.get(idx) {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            let src = o
                .properties
                .get("source")
                .map(|v| match v {
                    Value::String(s) => s.to_string(),
                    o => crate::keys::value_display_string(o),
                })
                .unwrap_or_default();
            let flags = o
                .properties
                .get("flags")
                .map(|v| match v {
                    Value::String(s) => s.to_string(),
                    o => crate::keys::value_display_string(o),
                })
                .unwrap_or_default();
            (src, flags)
        }
        Some(Value::String(s)) => split_regex_literal(s.as_ref()),
        Some(other) => split_regex_literal(&crate::keys::value_display_string(other)),
        None => (String::new(), String::new()),
    }
}

/// Pull pattern + flags out of a string in `/pat/flags` shape — what the
/// JS walker emits when it encounters a regex literal `/\d+/g`. The
/// pattern may contain escaped slashes (`\/`); we split on the LAST
/// unescaped `/`. Plain strings (no leading `/`) pass through as the
/// pattern with empty flags.
fn split_regex_literal(s: &str) -> (String, String) {
    if !s.starts_with('/') {
        return (s.to_string(), String::new());
    }
    // Find the LAST `/` not preceded by an odd number of backslashes.
    let bytes = s.as_bytes();
    let mut last = None;
    for (i, &b) in bytes.iter().enumerate().skip(1) {
        if b == b'/' {
            let mut bs = 0;
            let mut k = i;
            while k > 0 && bytes[k - 1] == b'\\' {
                bs += 1;
                k -= 1;
            }
            if bs % 2 == 0 {
                last = Some(i);
            }
        }
    }
    match last {
        Some(end) if end > 0 => {
            let pattern = s[1..end].to_string();
            let flags = s[end + 1..].to_string();
            (pattern, flags)
        }
        _ => (s.to_string(), String::new()),
    }
}

fn display_source(pattern: &str) -> String {
    pattern.replace('/', r#"\/"#)
}

fn special_pattern(pattern: &str, flags: &str) -> Option<SpecialPattern> {
    if flags.contains('v') && pattern == r"\p{RGI_Emoji}" {
        Some(SpecialPattern::RgiEmoji)
    } else {
        None
    }
}

fn empty_groups_object() -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new()))
}

fn range_to_value(start: usize, end: usize) -> Value {
    make_array(vec![Value::I32(start as i32), Value::I32(end as i32)])
}

fn exec_span_to_value(input: &str, start: usize, end: usize, include_indices: bool) -> Value {
    let mut match_obj = Object::new_array(vec![s_val(&input[start..end])]);
    match_obj
        .properties
        .reserve(if include_indices { 4 } else { 3 });
    match_obj
        .properties
        .insert("index".into(), Value::I32(start as i32));
    match_obj.properties.insert("input".into(), s_val(input));
    match_obj
        .properties
        .insert("groups".into(), empty_groups_object());
    if include_indices {
        let mut indices = Object::new_array(vec![range_to_value(start, end)]);
        indices.properties.reserve(1);
        indices
            .properties
            .insert("groups".into(), empty_groups_object());
        match_obj.properties.insert(
            "indices".into(),
            Value::Object(vybe_runtime::heap::alloc(indices)),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(match_obj))
}

fn special_match_at(input: &str, kind: SpecialPattern, start: usize) -> Option<usize> {
    if start > input.len() || !input.is_char_boundary(start) {
        return None;
    }
    let grapheme = input[start..].graphemes(true).next()?;
    match kind {
        SpecialPattern::RgiEmoji if is_rgi_emoji_sequence(grapheme) => Some(start + grapheme.len()),
        _ => None,
    }
}

fn special_find(input: &str, kind: SpecialPattern, start: usize) -> Option<(usize, usize)> {
    if start > input.len() || !input.is_char_boundary(start) {
        return None;
    }
    for (offset, grapheme) in input[start..].grapheme_indices(true) {
        let found = match kind {
            SpecialPattern::RgiEmoji => is_rgi_emoji_sequence(grapheme),
        };
        if found {
            let match_start = start + offset;
            return Some((match_start, match_start + grapheme.len()));
        }
    }
    None
}

fn special_find_all(input: &str, kind: SpecialPattern) -> Vec<(usize, usize)> {
    let mut out = Vec::new();
    for (start, grapheme) in input.grapheme_indices(true) {
        let found = match kind {
            SpecialPattern::RgiEmoji => is_rgi_emoji_sequence(grapheme),
        };
        if found {
            out.push((start, start + grapheme.len()));
        }
    }
    out
}

fn is_rgi_emoji_sequence(cluster: &str) -> bool {
    let chars: Vec<char> = cluster.chars().collect();
    if chars.is_empty() {
        return false;
    }

    let emoji = CodePointSetData::new::<Emoji>();
    let emoji_component = CodePointSetData::new::<EmojiComponent>();
    let emoji_modifier = CodePointSetData::new::<EmojiModifier>();
    let emoji_modifier_base = CodePointSetData::new::<EmojiModifierBase>();
    let emoji_presentation = CodePointSetData::new::<EmojiPresentation>();
    let regional_indicator = CodePointSetData::new::<RegionalIndicator>();

    let is_text_emoji_base = |ch: char| emoji.contains(ch) && !emoji_component.contains(ch);
    let is_simple_emoji_element = |segment: &str| {
        let segment_chars: Vec<char> = segment.chars().collect();
        match segment_chars.as_slice() {
            [ch] => emoji_presentation.contains(*ch),
            [ch, '\u{FE0F}'] => is_text_emoji_base(*ch),
            [base, modifier] => {
                emoji_modifier_base.contains(*base) && emoji_modifier.contains(*modifier)
            }
            [base, '\u{FE0F}', modifier] => {
                is_text_emoji_base(*base) && emoji_modifier.contains(*modifier)
            }
            _ => false,
        }
    };

    if matches!(chars.as_slice(), [first, second] if regional_indicator.contains(*first) && regional_indicator.contains(*second))
    {
        return true;
    }

    if matches!(chars.as_slice(), [base, '\u{20E3}'] if matches!(*base, '0'..='9' | '#' | '*')) {
        return true;
    }
    if matches!(chars.as_slice(), [base, '\u{FE0F}', '\u{20E3}'] if matches!(*base, '0'..='9' | '#' | '*'))
    {
        return true;
    }

    if chars.len() >= 3
        && chars.last() == Some(&'\u{E007F}')
        && is_text_emoji_base(chars[0])
        && chars[1..chars.len() - 1]
            .iter()
            .all(|ch| matches!(*ch, '\u{E0020}'..='\u{E007E}'))
    {
        return true;
    }

    if cluster.contains('\u{200D}') {
        let mut count = 0usize;
        for segment in cluster.split('\u{200D}') {
            if segment.is_empty() || !is_simple_emoji_element(segment) {
                return false;
            }
            count += 1;
        }
        return count >= 2;
    }

    is_simple_emoji_element(cluster)
}

fn regexp_exec_special(args: &[Value], input: &str, kind: SpecialPattern, flags: &str) -> Value {
    let is_global_or_sticky = flags.contains('g') || flags.contains('y');
    let is_sticky = flags.contains('y');
    let last_index = if is_global_or_sticky {
        args.first()
            .and_then(|v| match v {
                Value::Object(obj) => obj
                    .lock()
                    .unwrap()
                    .properties
                    .get("lastIndex")
                    .map(|v| v.as_i32()),
                _ => None,
            })
            .unwrap_or(0)
            .max(0) as usize
    } else {
        0
    };
    let search_start = last_index.min(input.len());
    let found = if is_sticky {
        special_match_at(input, kind, search_start).map(|end| (search_start, end))
    } else if is_global_or_sticky {
        special_find(input, kind, search_start)
    } else {
        special_find(input, kind, 0)
    };

    let (start, end) = match found {
        Some(found) => found,
        None => {
            if is_global_or_sticky {
                if let Some(Value::Object(obj)) = args.first() {
                    obj.lock()
                        .unwrap()
                        .properties
                        .insert("lastIndex".into(), Value::I32(0));
                }
            }
            return Value::Null;
        }
    };

    if is_global_or_sticky {
        if let Some(Value::Object(obj)) = args.first() {
            obj.lock()
                .unwrap()
                .properties
                .insert("lastIndex".into(), Value::I32(end as i32));
        }
    }
    exec_span_to_value(input, start, end, flags.contains('d'))
}

/// Compile a JS regexp using `regress`, filtering out wrapper-only flags.
fn compile(pattern: &str, flags: &str) -> Option<Arc<Regex>> {
    if special_pattern(pattern, flags).is_some() {
        return None;
    }
    if flags.contains('u') && flags.contains('v') {
        return None;
    }
    let normalized_pattern: Cow<'_, str> = if pattern.contains("(?P<") {
        Cow::Owned(pattern.replace("(?P<", "(?<"))
    } else {
        Cow::Borrowed(pattern)
    };
    let compile_flags = compile_flags_view(flags);
    cached_compile(normalized_pattern.as_ref(), compile_flags.as_ref())
}

fn cached_compile(pattern: &str, flags: &str) -> Option<Arc<Regex>> {
    if let Some(regex) = lookup_cached_regex(pattern, flags) {
        return Some(regex);
    }
    let regex = Arc::new(Regex::with_flags(pattern, flags).ok()?);
    Some(remember_compiled_regex(pattern, flags, regex))
}

fn lookup_cached_regex(pattern: &str, flags: &str) -> Option<Arc<Regex>> {
    REGEXP_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        if let Some((cached_pattern, cached_flags, regex)) = cache.front() {
            if cached_pattern.as_ref() == pattern && cached_flags.as_ref() == flags {
                return Some(Arc::clone(regex));
            }
        }
        if let Some(pos) = cache.iter().position(|(cached_pattern, cached_flags, _)| {
            cached_pattern.as_ref() == pattern && cached_flags.as_ref() == flags
        }) {
            if let Some(entry) = cache.remove(pos) {
                let regex = Arc::clone(&entry.2);
                cache.push_front(entry);
                return Some(regex);
            }
        }
        None
    })
}

fn remember_compiled_regex(pattern: &str, flags: &str, regex: Arc<Regex>) -> Arc<Regex> {
    REGEXP_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        if cache.len() == REGEXP_CACHE_LIMIT {
            cache.pop_back();
        }
        cache.push_front((Arc::from(pattern), Arc::from(flags), Arc::clone(&regex)));
    });
    regex
}

fn compile_flags_view(flags: &str) -> Cow<'_, str> {
    if flags
        .chars()
        .all(|c| matches!(c, 'i' | 'm' | 's' | 'u' | 'v'))
    {
        Cow::Borrowed(flags)
    } else {
        Cow::Owned(
            flags
                .chars()
                .filter(|c| matches!(c, 'i' | 'm' | 's' | 'u' | 'v'))
                .collect(),
        )
    }
}

fn match_group_value(m: &Match, input: &str, index: usize) -> Value {
    match m.group(index) {
        Some(range) => s_val(&input[range]),
        None => Value::Undefined,
    }
}

fn match_group_indices_value(m: &Match, index: usize, index_offset: usize) -> Value {
    match m.group(index) {
        Some(range) => range_to_value(index_offset + range.start, index_offset + range.end),
        None => Value::Undefined,
    }
}

fn named_groups_object(m: &Match, input: &str) -> Value {
    let named_count = m.named_groups().count();
    let mut groups = Object::new();
    groups
        .properties
        .reserve(named_count + usize::from(named_count != 0));
    let mut group_order: Vec<Value> = Vec::with_capacity(named_count);
    for (name, _) in m.named_groups() {
        let value = match m.named_group(name) {
            Some(range) => s_val(&input[range]),
            None => Value::Undefined,
        };
        groups.properties.insert(name.to_string(), value);
        group_order.push(s_val(name));
    }
    if !group_order.is_empty() {
        groups.properties.insert(
            "__keys".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(group_order))),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(groups))
}

fn named_groups_indices_object(m: &Match, index_offset: usize) -> Value {
    let named_count = m.named_groups().count();
    let mut groups = Object::new();
    groups
        .properties
        .reserve(named_count + usize::from(named_count != 0));
    let mut group_order: Vec<Value> = Vec::with_capacity(named_count);
    for (name, _) in m.named_groups() {
        let value = match m.named_group(name) {
            Some(range) => range_to_value(index_offset + range.start, index_offset + range.end),
            None => Value::Undefined,
        };
        groups.properties.insert(name.to_string(), value);
        group_order.push(s_val(name));
    }
    if !group_order.is_empty() {
        groups.properties.insert(
            "__keys".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(group_order))),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(groups))
}

fn exec_match_to_value(
    m: &Match,
    input: &str,
    index_offset: usize,
    include_indices: bool,
) -> Value {
    let mut elems: Vec<Value> = Vec::with_capacity(m.captures.len() + 1);
    for i in 0..=m.captures.len() {
        elems.push(match_group_value(m, input, i));
    }
    let mut match_obj = Object::new_array(elems);
    match_obj
        .properties
        .reserve(if include_indices { 4 } else { 3 });
    match_obj.properties.insert(
        "index".into(),
        Value::I32((index_offset + m.start()) as i32),
    );
    match_obj.properties.insert("input".into(), s_val(input));
    match_obj
        .properties
        .insert("groups".into(), named_groups_object(m, input));
    if include_indices {
        let mut indices: Vec<Value> = Vec::with_capacity(m.captures.len() + 1);
        for i in 0..=m.captures.len() {
            indices.push(match_group_indices_value(m, i, index_offset));
        }
        let mut indices_obj = Object::new_array(indices);
        indices_obj.properties.reserve(1);
        indices_obj.properties.insert(
            "groups".into(),
            named_groups_indices_object(m, index_offset),
        );
        match_obj.properties.insert(
            "indices".into(),
            Value::Object(vybe_runtime::heap::alloc(indices_obj)),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(match_obj))
}

fn make_array(elements: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(elements)))
}

fn s_val(s: &str) -> Value {
    crate::keys::string_value(s)
}

#[inline]
fn s_owned(s: String) -> Value {
    crate::keys::owned_string_value(s)
}

#[inline]
fn s_cow(s: Cow<'_, str>) -> Value {
    match s {
        Cow::Borrowed(text) => s_val(text),
        Cow::Owned(text) => s_owned(text),
    }
}

fn validate_flags(flags: &str) -> Result<(), String> {
    let mut seen = 0u8;
    for flag in flags.chars() {
        let bit = match flag {
            'd' => 1,
            'g' => 2,
            'i' => 4,
            'm' => 8,
            's' => 16,
            'u' => 32,
            'v' => 64,
            'y' => 128,
            _ => return Err(format!("Invalid regular expression flag '{}'", flag)),
        };
        if seen & bit != 0 {
            return Err(format!("Duplicate regular expression flag '{}'", flag));
        }
        seen |= bit;
    }
    if flags.contains('u') && flags.contains('v') {
        return Err("Regular expression flags 'u' and 'v' cannot be combined".into());
    }
    Ok(())
}

fn validate_pattern(pattern: &str, flags: &str) -> Result<(), String> {
    if special_pattern(pattern, flags).is_some() {
        return Ok(());
    }
    let normalized_pattern: Cow<'_, str> = if pattern.contains("(?P<") {
        Cow::Owned(pattern.replace("(?P<", "(?<"))
    } else {
        Cow::Borrowed(pattern)
    };
    let compile_flags = compile_flags_view(flags);
    if lookup_cached_regex(normalized_pattern.as_ref(), compile_flags.as_ref()).is_some() {
        return Ok(());
    }
    Regex::with_flags(normalized_pattern.as_ref(), compile_flags.as_ref())
        .map(|regex| {
            remember_compiled_regex(normalized_pattern.as_ref(), compile_flags.as_ref(), Arc::new(regex));
        })
        .map_err(|err| format!("Invalid regular expression: {}", err))
}

#[cfg(test)]
mod speedup_tests {
    use super::*;

    #[test]
    fn constructor_validation_populates_execution_cache() {
        REGEXP_CACHE.with(|cache| cache.borrow_mut().clear());
        validate_pattern("a(b+)", "gi").unwrap();
        let validated = lookup_cached_regex("a(b+)", "i").expect("validation must cache compiled code");
        let executed = compile("a(b+)", "gi").unwrap();
        assert!(Arc::ptr_eq(&validated, &executed));
        validate_pattern("a(b+)", "gi").unwrap();
        assert!(Arc::ptr_eq(&validated, &lookup_cached_regex("a(b+)", "i").unwrap()));
    }

    #[test]
    fn cache_stays_bounded_and_reused_patterns_are_promoted() {
        REGEXP_CACHE.with(|cache| cache.borrow_mut().clear());
        for index in 0..REGEXP_CACHE_LIMIT {
            cached_compile(&format!("pattern_{index}"), "").unwrap();
        }
        let oldest = lookup_cached_regex("pattern_0", "").unwrap();
        cached_compile("additional_pattern", "").unwrap();
        assert!(lookup_cached_regex("pattern_1", "").is_none());
        assert!(Arc::ptr_eq(&oldest, &lookup_cached_regex("pattern_0", "").unwrap()));
        REGEXP_CACHE.with(|cache| assert_eq!(cache.borrow().len(), REGEXP_CACHE_LIMIT));
    }

    #[test]
    fn flag_errors_and_invalid_patterns_are_preserved() {
        assert!(validate_flags("dgimsy").is_ok());
        assert!(validate_flags("v").is_ok());
        assert_eq!(validate_flags("ii").unwrap_err(), "Duplicate regular expression flag 'i'");
        assert_eq!(validate_flags("x").unwrap_err(), "Invalid regular expression flag 'x'");
        assert_eq!(validate_flags("uv").unwrap_err(), "Regular expression flags 'u' and 'v' cannot be combined");
        assert!(validate_pattern("[", "").unwrap_err().starts_with("Invalid regular expression: "));
        assert!(lookup_cached_regex("[", "").is_none());
    }

    #[test]
    fn streamed_replacements_preserve_expansion_and_match_order() {
        let captures = Regex::with_flags("(a)(b)?", "").unwrap();
        assert_eq!(apply_string_replacement("aba", &captures, true, "<$1:$2>"), "<a:b><a:>");
        let single = Regex::with_flags("a", "").unwrap();
        assert_eq!(apply_string_replacement("cat", &single, false, "$$:$&:$`:$'"), "c$:a:c:tt");
        assert_eq!(apply_string_replacement("aa", &single, false, "x"), "xa");
        assert_eq!(apply_string_replacement("aa", &single, true, "x"), "xx");
        assert_eq!(apply_string_replacement("aa", &single, true, "\u{e9}$&"), "\u{e9}a\u{e9}a");
        assert!(matches!(apply_string_replacement("xyz", &single, true, "$&"), Cow::Borrowed("xyz")));
        let named = Regex::with_flags("(?<letter>a)", "").unwrap();
        assert_eq!(apply_string_replacement("aa", &named, true, "[$<letter>]"), "[a][a]");
        assert_eq!(apply_string_replacement("aa", &named, true, "$<missing>"), "$<missing>$<missing>");
        let empty = Regex::with_flags("", "").unwrap();
        assert_eq!(apply_string_replacement("ab", &empty, true, "-"), "-a-b-");
        assert_eq!(apply_string_replacement(&"a".repeat(1024), &single, true, "xyz"), "xyz".repeat(1024));
    }

    #[test]
    #[ignore = "explicit native microbenchmark; timings are not a conformance gate"]
    fn native_speedup_microbenchmark() {
        use std::hint::black_box;
        use std::time::Instant;

        const ITERATIONS: usize = 1000;
        let pattern = r"(?<key>[A-Za-z_][A-Za-z0-9_]*)\s*=\s*(\d+)";
        let flags = "im";
        validate_pattern(pattern, flags).unwrap();
        let start = Instant::now();
        for _ in 0..ITERATIONS {
            Regex::with_flags(black_box(pattern), black_box(flags))
                .map(|_| ())
                .unwrap();
        }
        let compile_each = start.elapsed();
        let start = Instant::now();
        for _ in 0..ITERATIONS {
            validate_pattern(black_box(pattern), black_box(flags)).unwrap();
        }
        let reused = start.elapsed();
        println!("debug native validation, n={ITERATIONS}: compile-each={compile_each:?}, warm-reuse={reused:?}");

        // Reconstruct the prior per-match template and expanded-string storage
        // using the same expansion code, so the compared outputs are identical.
        let legacy_replace = |input: &str, regex: &Regex, replacement: &str| {
            let mut out = String::with_capacity(input.len());
            let mut end = 0;
            for matched in regex.find_iter(input) {
                out.push_str(&input[end..matched.start()]);
                let chars: Vec<char> = replacement.chars().collect();
                let mut expanded = String::new();
                expand_js_replacement_into(&chars, input, &matched, &mut expanded);
                out.push_str(&expanded);
                end = matched.end();
            }
            out.push_str(&input[end..]);
            out
        };
        let regex = Regex::with_flags("a", "").unwrap();
        let input = "a".repeat(256);
        for replacement in ["xyz", "<$&>"] {
            assert_eq!(apply_string_replacement(&input, &regex, true, replacement), legacy_replace(&input, &regex, replacement));
            let start = Instant::now();
            for _ in 0..ITERATIONS {
                black_box(legacy_replace(black_box(&input), &regex, black_box(replacement)));
            }
            let per_match_buffers = start.elapsed();
            let start = Instant::now();
            for _ in 0..ITERATIONS {
                black_box(apply_string_replacement(black_box(&input), &regex, true, black_box(replacement)));
            }
            let streamed = start.elapsed();
            println!("debug native replacement, n={ITERATIONS}, matches/call=256, template={replacement:?}: per-match-buffers={per_match_buffers:?}, streamed={streamed:?}");
        }
    }
}

fn throw_syntax_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::error::new_error(ctx, "SyntaxError", message));
    Value::Null
}

fn throw_type_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::error::new_error(ctx, "TypeError", message));
    Value::Null
}

fn lookup_symbol_method(target: &Value, key: &str) -> Option<Value> {
    let Value::Object(obj) = target else {
        return None;
    };
    let mut current = Some(obj.clone());
    for _ in 0..100 {
        let Some(cur) = current else {
            break;
        };
        let (prop, next_proto) = {
            let o = cur.lock().unwrap();
            (
                o.properties.get(key).cloned(),
                match o.properties.get("__proto__").cloned() {
                    Some(Value::Object(proto)) => Some(proto),
                    _ => None,
                },
            )
        };
        if let Some(value) = prop {
            if !matches!(value, Value::Null | Value::Undefined) {
                return Some(value);
            }
        }
        current = next_proto;
    }
    None
}

pub fn register(vm: &mut VM) {
    register_constructor(vm);
    register_prototype(vm);
    register_string_methods(vm);
}

// ── Constructor ───────────────────────────────────────────────────────

// ── RegExp.prototype ─────────────────────────────────────────────────

fn register_constructor(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:regexp",
        "new",
        Box::new(|ctx, args| {
            let (pattern, default_flags) = extract_pattern(args, 0);
            // Explicit flags arg overrides any flags inherited from a
            // RegExp first arg.
            let flags = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(Value::Undefined) | None => default_flags,
                Some(other) => crate::keys::value_display_string(other),
            };
            if let Err(message) = validate_flags(&flags) {
                return throw_syntax_error(ctx, &message);
            }
            if let Err(message) = validate_pattern(&pattern, &flags) {
                return throw_syntax_error(ctx, &message);
            }
            let mut obj = Object::new();
            obj.properties.reserve(13);
            obj.properties
                .insert("source".into(), s_val(&display_source(&pattern)));
            obj.properties.insert("flags".into(), s_val(&flags));
            obj.properties
                .insert("global".into(), Value::Bool(flags.contains('g')));
            obj.properties
                .insert("ignoreCase".into(), Value::Bool(flags.contains('i')));
            obj.properties
                .insert("multiline".into(), Value::Bool(flags.contains('m')));
            obj.properties
                .insert("dotAll".into(), Value::Bool(flags.contains('s')));
            obj.properties
                .insert("unicode".into(), Value::Bool(flags.contains('u')));
            obj.properties
                .insert("unicodeSets".into(), Value::Bool(flags.contains('v')));
            obj.properties
                .insert("sticky".into(), Value::Bool(flags.contains('y')));
            obj.properties
                .insert("hasIndices".into(), Value::Bool(flags.contains('d')));
            obj.properties.insert("lastIndex".into(), Value::I32(0));
            // __type lets cross-language `instanceof RegExp` work via the
            // type registry; matches the pattern used by Map/Set/etc.
            obj.properties
                .insert("__type".into(), crate::keys::string_value(REGEXP_TYPE));
            obj.properties
                .insert("__proto__".into(), shared_regexp_prototype());
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // newWithFlags(pattern, flags) — explicit flags alias.
    vm.register_host_fn(
        "ecma:regexp",
        "newWithFlags",
        Box::new(|ctx, args| {
            let (pattern, _) = extract_pattern(args, 0);
            let flags = match args.get(1) {
                Some(Value::String(s)) => s.as_ref().to_owned(),
                Some(v) => crate::keys::value_display_string(v),
                None => String::new(),
            };
            if let Err(message) = validate_flags(&flags) {
                return throw_syntax_error(ctx, &message);
            }
            if let Err(message) = validate_pattern(&pattern, &flags) {
                return throw_syntax_error(ctx, &message);
            }
            let mut obj = Object::new();
            obj.properties.reserve(13);
            obj.properties
                .insert("source".into(), s_val(&display_source(&pattern)));
            obj.properties.insert("flags".into(), s_val(&flags));
            obj.properties
                .insert("global".into(), Value::Bool(flags.contains('g')));
            obj.properties
                .insert("ignoreCase".into(), Value::Bool(flags.contains('i')));
            obj.properties
                .insert("multiline".into(), Value::Bool(flags.contains('m')));
            obj.properties
                .insert("dotAll".into(), Value::Bool(flags.contains('s')));
            obj.properties
                .insert("unicode".into(), Value::Bool(flags.contains('u')));
            obj.properties
                .insert("unicodeSets".into(), Value::Bool(flags.contains('v')));
            obj.properties
                .insert("sticky".into(), Value::Bool(flags.contains('y')));
            obj.properties
                .insert("hasIndices".into(), Value::Bool(flags.contains('d')));
            obj.properties.insert("lastIndex".into(), Value::I32(0));
            obj.properties
                .insert("__type".into(), crate::keys::string_value(REGEXP_TYPE));
            obj.properties
                .insert("__proto__".into(), shared_regexp_prototype());
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // escape(str) — ES2025 §22.2.2.1. Escape all special regex chars.
    vm.register_host_fn(
        "ecma:regexp",
        "escape",
        Box::new(|_ctx, args| {
            let s: Cow<'_, str> = match args.first() {
                Some(Value::String(s)) => Cow::Borrowed(s.as_ref()),
                Some(v) => crate::keys::value_display_cow(v),
                None => return Value::Undefined,
            };
            let mut escaped = String::with_capacity(s.len());
            for c in s.chars() {
                if c.is_alphanumeric() || c == '_' {
                    escaped.push(c);
                } else {
                    escaped.push('\\');
                    escaped.push(c);
                }
            }
            s_owned(escaped)
        }),
    );
}

// ── RegExp.prototype ─────────────────────────────────────────────────

fn register_prototype(vm: &mut VM) {
    // `regex.test(str)` — ECMA-262 §22.2.5.15. True iff pattern matches
    // anywhere in str. Receiver is `args[0]` per Component-Model
    // `[method]` convention.
    vm.register_host_fn(
        "ecma:regexp",
        "test",
        Box::new(|ctx, args| regexp_test(ctx, args)),
    );

    // `regex.exec(str)` — ECMA-262 §22.2.5.2. Returns a match Array
    // `[full, g1, g2, ..., index, input, groups]` or null.
    //
    // Spec layout: the array's numeric elements are full + capture groups,
    // with `.index`, `.input`, and `.groups` set as own properties on the
    // array. We materialize all of these so `match[0]`, `match.index`,
    // and `match.groups.name` all work.
    vm.register_host_fn(
        "ecma:regexp",
        "exec",
        Box::new(|ctx, args| regexp_exec(ctx, args)),
    );

    // `regex.toString()` — ECMA-262 §22.2.5.17. Returns "/source/flags".
    vm.register_host_fn(
        "ecma:regexp",
        "toString",
        Box::new(|_ctx, args| regexp_to_string(args)),
    );
}

pub fn dispatch_regexp_method(
    ctx: &mut HostContext,
    method: &str,
    args: &[Value],
) -> Option<Value> {
    match method {
        "test" => Some(regexp_test(ctx, args)),
        "exec" => Some(regexp_exec(ctx, args)),
        "toString" => Some(regexp_to_string(args)),
        _ => None,
    }
}

pub fn dispatch_regexp_string_method(
    ctx: &mut HostContext,
    method: &str,
    args: &[Value],
) -> Option<Value> {
    match method {
        "match" => Some(regexp_string_match(ctx, args)),
        "matchAll" => Some(regexp_string_match_all(ctx, args)),
        "search" => Some(regexp_string_search(ctx, args)),
        "replace" => Some(regexp_string_replace(ctx, args)),
        "replaceAll" => Some(regexp_string_replace_all(args)),
        "split" => Some(regexp_string_split(args)),
        _ => None,
    }
}

fn regexp_test(ctx: &mut HostContext, args: &[Value]) -> Value {
    Value::Bool(!matches!(regexp_exec(ctx, args), Value::Null))
}

fn regexp_exec(ctx: &mut HostContext, args: &[Value]) -> Value {
    let (pattern, flags) = extract_pattern(args, 0);
    // §22.2.7.2 RegExpBuiltinExec step 1: ToString(argument). A non-string
    // argument is coerced through its `toString`/`valueOf` (ToPrimitive with a
    // string hint) rather than the raw `[object]` Display.
    let input: Cow<'_, str> = match args.get(1) {
        Some(Value::String(s)) => Cow::Borrowed(s.as_ref()),
        Some(other) => {
            let primitive = crate::value::to_primitive(ctx, other, "string");
            Cow::Owned(crate::keys::value_display_string(&primitive))
        }
        None => Cow::Borrowed(""),
    };
    if let Some(kind) = special_pattern(&pattern, &flags) {
        return regexp_exec_special(args, &input, kind, &flags);
    }
    let re = match compile(&pattern, &flags) {
        Some(re) => re,
        None => return Value::Null,
    };
    let is_global_or_sticky = flags.contains('g') || flags.contains('y');
    let is_sticky = flags.contains('y');
    let last_index = if is_global_or_sticky {
        args.first()
            .and_then(|v| match v {
                Value::Object(obj) => obj
                    .lock()
                    .unwrap()
                    .properties
                    .get("lastIndex")
                    .map(|v| v.as_i32()),
                _ => None,
            })
            .unwrap_or(0)
            .max(0) as usize
    } else {
        0
    };
    let search_start = last_index.min(input.len());
    let found = if is_global_or_sticky {
        re.find_from(&input, search_start).next()
    } else {
        re.find(&input)
    };
    let m = match found {
        Some(m) if !is_sticky || m.start() == search_start => m,
        Some(_) => {
            if is_global_or_sticky {
                if let Some(Value::Object(obj)) = args.first() {
                    obj.lock()
                        .unwrap()
                        .properties
                        .insert("lastIndex".into(), Value::I32(0));
                }
            }
            return Value::Null;
        }
        None => {
            if is_global_or_sticky {
                if let Some(Value::Object(obj)) = args.first() {
                    obj.lock()
                        .unwrap()
                        .properties
                        .insert("lastIndex".into(), Value::I32(0));
                }
            }
            return Value::Null;
        }
    };
    if is_global_or_sticky {
        let new_idx = m.end() as i32;
        if let Some(Value::Object(obj)) = args.first() {
            obj.lock()
                .unwrap()
                .properties
                .insert("lastIndex".into(), Value::I32(new_idx));
        }
    }
    exec_match_to_value(&m, &input, 0, flags.contains('d'))
}

fn regexp_to_string(args: &[Value]) -> Value {
    let (pattern, flags) = extract_pattern(args, 0);
    s_owned(format!("/{}/{}", pattern, flags))
}

// ── String.prototype regex methods ───────────────────────────────────
//
// These take a string receiver + RegExp argument. Live under
// `ecma:regexp` (rather than `ecma:string`) because the regexp compiler
// is the load-bearing dependency — keeping all regex-using ops in one
// place makes flag handling consistent and keeps the engine swap local.

fn register_string_methods(vm: &mut VM) {
    // `str.match(regex)` — §22.1.3.13. Without `g`: same as
    // `regex.exec(str)` (single match Array with groups). With `g`:
    // Array of full-match strings only (no groups).
    vm.register_host_fn(
        "ecma:regexp",
        "match",
        Box::new(|ctx, args| regexp_string_match(ctx, args)),
    );

    // `str.matchAll(regex)` — §22.1.3.14. Spec returns an iterator;
    // MVP returns an Array of match Arrays (each shaped like exec's
    // result). Iterator semantics layer on top once iterator protocol
    // dispatch lands.
    vm.register_host_fn(
        "ecma:regexp",
        "matchAll",
        Box::new(|ctx, args| regexp_string_match_all(ctx, args)),
    );

    // `str.search(regex)` — §22.1.3.16. Returns index of first match
    // or -1.
    vm.register_host_fn(
        "ecma:regexp",
        "search",
        Box::new(|ctx, args| regexp_string_search(ctx, args)),
    );

    // `str.replace(regex, replacement)` — §22.1.3.18. Replaces first
    // match (or all if `g` flag is set, per spec). Replacement is either
    // a string (with $1/$2/$<name> capture refs) or a function called
    // with (match, ...captures, offset, input). The function form needs
    // VM callback dispatch via `ctx.invoke`.
    vm.register_host_fn(
        "ecma:regexp",
        "replace",
        Box::new(|ctx, args| regexp_string_replace(ctx, args)),
    );

    // `str.replaceAll(regex, replacement)` — §22.1.3.19. With a RegExp,
    // requires the `g` flag (otherwise spec throws TypeError); we just
    // replace-all unconditionally for simplicity.
    vm.register_host_fn(
        "ecma:regexp",
        "replaceAll",
        Box::new(|_ctx, args| regexp_string_replace_all(args)),
    );

    // `str.split(regex, limit?)` — §22.1.3.20. Splits on regex matches.
    vm.register_host_fn(
        "ecma:regexp",
        "split",
        Box::new(|_ctx, args| regexp_string_split(args)),
    );
}

fn regexp_string_match(ctx: &mut HostContext, args: &[Value]) -> Value {
    with_s_arg(args, 0, |input| {
        if let Some(method) = args
            .get(1)
            .and_then(|value| lookup_symbol_method(value, "symbolmatch"))
        {
            return ctx.invoke(&method, &[s_val(input)]);
        }
        let (pattern, flags) = extract_pattern(args, 1);
        if let Some(kind) = special_pattern(&pattern, &flags) {
            if flags.contains('g') {
                let spans = special_find_all(input, kind);
                let mut matches = Vec::with_capacity(spans.len());
                for (start, end) in spans {
                    matches.push(s_val(&input[start..end]));
                }
                return if matches.is_empty() {
                    Value::Null
                } else {
                    make_array(matches)
                };
            }
            return match special_find(input, kind, 0) {
                Some((start, end)) => exec_span_to_value(input, start, end, flags.contains('d')),
                None => Value::Null,
            };
        }
        let re = match compile(&pattern, &flags) {
            Some(re) => re,
            None => return Value::Null,
        };
        if flags.contains('g') {
            let mut matches = Vec::new();
            for m in re.find_iter(input) {
                matches.push(s_val(m.as_str(input)));
            }
            if matches.is_empty() {
                Value::Null
            } else {
                make_array(matches)
            }
        } else {
            match re.find(input) {
                Some(m) => exec_match_to_value(&m, input, 0, flags.contains('d')),
                None => Value::Null,
            }
        }
    })
}

fn regexp_string_match_all(ctx: &mut HostContext, args: &[Value]) -> Value {
    with_s_arg(args, 0, |input| {
        let (pattern, flags) = extract_pattern(args, 1);
        if regex_like_arg(args.get(1)) && !flags.contains('g') {
            return throw_type_error(
                ctx,
                "String.prototype.matchAll called with a non-global RegExp argument",
            );
        }
        if let Some(kind) = special_pattern(&pattern, &flags) {
            let spans = special_find_all(input, kind);
            let mut matches = Vec::with_capacity(spans.len());
            for (start, end) in spans {
                matches.push(exec_span_to_value(input, start, end, flags.contains('d')));
            }
            return make_array(matches);
        }
        let re = match compile(&pattern, &flags) {
            Some(re) => re,
            None => return make_array(Vec::new()),
        };
        let mut out = Vec::new();
        for m in re.find_iter(input) {
            out.push(exec_match_to_value(&m, input, 0, flags.contains('d')));
        }
        make_array(out)
    })
}

fn regexp_string_search(ctx: &mut HostContext, args: &[Value]) -> Value {
    with_s_arg(args, 0, |input| {
        if let Some(method) = args
            .get(1)
            .and_then(|value| lookup_symbol_method(value, "symbolsearch"))
        {
            return ctx.invoke(&method, &[s_val(input)]);
        }
        let (pattern, flags) = extract_pattern(args, 1);
        if let Some(kind) = special_pattern(&pattern, &flags) {
            return match special_find(input, kind, 0) {
                Some((start, _)) => Value::I32(start as i32),
                None => Value::I32(-1),
            };
        }
        match compile(&pattern, &flags) {
            Some(re) => match re.find(input) {
                Some(m) => Value::I32(m.start() as i32),
                None => Value::I32(-1),
            },
            None => Value::I32(-1),
        }
    })
}

fn regexp_string_replace(ctx: &mut HostContext, args: &[Value]) -> Value {
    with_s_arg(args, 0, |input| {
        let (pattern, flags) = extract_pattern(args, 1);
        let re = match compile(&pattern, &flags) {
            Some(re) => re,
            None => return s_val(input),
        };
        let global = flags.contains('g');
        let replacement_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
        let is_callable = matches!(&replacement_arg, Value::Object(o)
            if matches!(o.lock().unwrap().kind,
                vybe_runtime::value::ObjectKind::Function(_)
                | vybe_runtime::value::ObjectKind::HostFunction(_)));
        if is_callable {
            let prepared_replacement = crate::function::prepare_bound_callback(&replacement_arg);
            let mut out = String::with_capacity(input.len());
            let mut last_end = 0;
            let mut apply_match = |m: Match| {
                out.push_str(&input[last_end..m.start()]);
                let has_named_groups = m.named_groups().next().is_some();
                let ret = if m.captures.len() == 0 && !has_named_groups {
                    let cb_args = [
                        match_group_value(&m, input, 0),
                        Value::I32(m.start() as i32),
                        s_val(input),
                    ];
                    if let Some(prepared) = &prepared_replacement {
                        crate::function::invoke_prepared_bound_callback(ctx, prepared, &cb_args)
                    } else {
                        ctx.invoke(&replacement_arg, &cb_args)
                    }
                } else {
                    let total_len = m.captures.len() + 3 + usize::from(has_named_groups);
                    if total_len <= REGEXP_REPLACE_INLINE_ARG_LIMIT {
                        let mut cb_args: [Value; REGEXP_REPLACE_INLINE_ARG_LIMIT] =
                            std::array::from_fn(|_| Value::Undefined);
                        let mut out_index = 0usize;
                        for i in 0..=m.captures.len() {
                            cb_args[out_index] = match_group_value(&m, input, i);
                            out_index += 1;
                        }
                        cb_args[out_index] = Value::I32(m.start() as i32);
                        out_index += 1;
                        cb_args[out_index] = s_val(input);
                        out_index += 1;
                        if has_named_groups {
                            cb_args[out_index] = named_groups_object(&m, input);
                            out_index += 1;
                        }
                        if let Some(prepared) = &prepared_replacement {
                            crate::function::invoke_prepared_bound_callback(
                                ctx,
                                prepared,
                                &cb_args[..out_index],
                            )
                        } else {
                            ctx.invoke(&replacement_arg, &cb_args[..out_index])
                        }
                    } else {
                        let mut cb_args: Vec<Value> = Vec::with_capacity(total_len);
                        for i in 0..=m.captures.len() {
                            cb_args.push(match_group_value(&m, input, i));
                        }
                        cb_args.push(Value::I32(m.start() as i32));
                        cb_args.push(s_val(input));
                        if has_named_groups {
                            cb_args.push(named_groups_object(&m, input));
                        }
                        if let Some(prepared) = &prepared_replacement {
                            crate::function::invoke_prepared_bound_callback(ctx, prepared, &cb_args)
                        } else {
                            ctx.invoke(&replacement_arg, &cb_args)
                        }
                    }
                };
                match ret {
                    Value::String(s) => out.push_str(s.as_ref()),
                    other => {
                        let _ = write!(out, "{}", other);
                    }
                }
                last_end = m.end();
            };
            if global {
                for m in re.find_iter(input) {
                    apply_match(m);
                }
            } else if let Some(m) = re.find(input) {
                apply_match(m);
            }
            out.push_str(&input[last_end..]);
            return s_owned(out);
        }
        with_s_arg(args, 2, |replacement| {
            s_cow(apply_string_replacement(input, &re, global, replacement))
        })
    })
}

fn regexp_string_replace_all(args: &[Value]) -> Value {
    with_two_s_args(args, 0, 2, |input, replacement| {
        let (pattern, flags) = extract_pattern(args, 1);
        match compile(&pattern, &flags) {
            Some(re) => s_cow(apply_string_replacement(input, &re, true, replacement)),
            None => s_val(input),
        }
    })
}

fn regexp_string_split(args: &[Value]) -> Value {
    with_s_arg(args, 0, |input| {
        let (pattern, flags) = extract_pattern(args, 1);
        let limit = args.get(2).map(|v| v.as_i32().max(0) as usize);
        if matches!(limit, Some(0)) {
            return make_array(Vec::new());
        }
        let max_parts = limit.unwrap_or(usize::MAX);
        if pattern.is_empty() || pattern == "(?:)" {
            let mut parts = Vec::with_capacity(input.len().min(max_parts));
            for ch in input.chars().take(max_parts) {
                parts.push(char_value(ch));
            }
            return make_array(parts);
        }
        match compile(&pattern, &flags) {
            Some(re) => {
                let mut parts: Vec<Value> =
                    Vec::with_capacity(max_parts.min(16).min(input.len().saturating_add(1)));
                let mut last_end = 0;
                for m in re.find_iter(input) {
                    if parts.len() >= max_parts {
                        break;
                    }
                    parts.push(s_val(&input[last_end..m.start()]));
                    for index in 1..=m.captures.len() {
                        if parts.len() >= max_parts {
                            break;
                        }
                        parts.push(match_group_value(&m, input, index));
                    }
                    last_end = m.end();
                }
                if parts.len() < max_parts {
                    parts.push(s_val(&input[last_end..]));
                }
                make_array(parts)
            }
            None => make_array(vec![s_val(input)]),
        }
    })
}

fn regex_like_arg(arg: Option<&Value>) -> bool {
    let Some(Value::Object(obj)) = arg else {
        return false;
    };
    let guard = obj.lock().unwrap();
    matches!(guard.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == REGEXP_TYPE)
}

fn apply_string_replacement<'a>(
    input: &'a str,
    re: &Regex,
    global: bool,
    replacement: &str,
) -> Cow<'a, str> {
    let mut matches = re.find_iter(input);
    let Some(first) = matches.next() else {
        return Cow::Borrowed(input);
    };
    let mut out = String::with_capacity(input.len());
    let mut last_end = 0;
    let template = replacement.contains('$').then(|| replacement.chars().collect::<Vec<_>>());
    let limit = if global { usize::MAX } else { 1 };
    for m in std::iter::once(first).chain(matches).take(limit) {
        out.push_str(&input[last_end..m.start()]);
        if let Some(template) = &template {
            expand_js_replacement_into(template, input, &m, &mut out);
        } else {
            out.push_str(replacement);
        }
        last_end = m.end();
    }
    out.push_str(&input[last_end..]);
    Cow::Owned(out)
}

fn expand_js_replacement_into(chars: &[char], input: &str, m: &Match, out: &mut String) {
    let whole = m.group(0).map(|range| &input[range]).unwrap_or("");
    let prefix = &input[..m.start()];
    let suffix = &input[m.end()..];
    let mut index = 0;
    while index < chars.len() {
        if chars[index] != '$' || index + 1 >= chars.len() {
            out.push(chars[index]);
            index += 1;
            continue;
        }
        match chars[index + 1] {
            '$' => {
                out.push('$');
                index += 2;
            }
            '&' => {
                out.push_str(whole);
                index += 2;
            }
            '`' => {
                out.push_str(prefix);
                index += 2;
            }
            '\'' => {
                out.push_str(suffix);
                index += 2;
            }
            digit if digit.is_ascii_digit() => {
                let first = digit.to_digit(10).unwrap_or(0) as usize;
                let mut group_index = first;
                let mut consumed = 2;
                if index + 2 < chars.len() && chars[index + 2].is_ascii_digit() {
                    let second = chars[index + 2].to_digit(10).unwrap_or(0) as usize;
                    let candidate = first * 10 + second;
                    if m.group(candidate).is_some() {
                        group_index = candidate;
                        consumed = 3;
                    }
                }
                if let Some(range) = m.group(group_index) {
                    out.push_str(&input[range]);
                }
                index += consumed;
            }
            '<' => {
                if let Some(close_offset) = chars[index + 2..].iter().position(|ch| *ch == '>') {
                    let name: String = chars[index + 2..index + 2 + close_offset].iter().collect();
                    let named_group_exists =
                        m.named_groups().any(|(group_name, _)| group_name == name);
                    if named_group_exists {
                        if let Some(range) = m.named_group(&name) {
                            out.push_str(&input[range]);
                        }
                        index += close_offset + 3;
                    } else {
                        out.push('$');
                        index += 1;
                    }
                } else {
                    out.push('$');
                    index += 1;
                }
            }
            _ => {
                out.push('$');
                index += 1;
            }
        }
    }
}
