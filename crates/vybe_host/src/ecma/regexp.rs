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
use std::sync::{Arc, Mutex};
use unicode_segmentation::UnicodeSegmentation;
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{HostContext, VM};

const REGEXP_TYPE: &str = "RegExp";

static REGEXP_PROTOTYPE: std::sync::OnceLock<Arc<Mutex<Object>>> = std::sync::OnceLock::new();

/// %RegExp.prototype% — process-global singleton (same pattern as the
/// other builtin prototypes). Instances link to it via `__proto__`, so
/// `Object.getPrototypeOf(/a/) === RegExp.prototype` and
/// `RegExp.prototype.isPrototypeOf(/a/)` hold (§22.2.6).
pub(crate) fn shared_regexp_prototype() -> Value {
    Value::Object(
        REGEXP_PROTOTYPE
            .get_or_init(|| {
                let mut proto = Object::new();
                proto.properties.insert(
                    "__proto__".into(),
                    crate::ecma::object::shared_object_prototype(),
                );
                Arc::new(Mutex::new(proto))
            })
            .clone(),
    )
}

#[derive(Clone, Copy)]
enum SpecialPattern {
    RgiEmoji,
}

fn s_arg(args: &[Value], idx: usize) -> String {
    match args.get(idx) {
        Some(Value::String(s)) => s.to_string(),
        Some(other) => format!("{}", other),
        None => String::new(),
    }
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
                    o => format!("{}", o),
                })
                .unwrap_or_default();
            let flags = o
                .properties
                .get("flags")
                .map(|v| match v {
                    Value::String(s) => s.to_string(),
                    o => format!("{}", o),
                })
                .unwrap_or_default();
            (src, flags)
        }
        Some(Value::String(s)) => split_regex_literal(s.as_ref()),
        Some(other) => split_regex_literal(&format!("{}", other)),
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
    Value::Object(Arc::new(Mutex::new(Object::new())))
}

fn range_to_value(start: usize, end: usize) -> Value {
    make_array(vec![Value::I32(start as i32), Value::I32(end as i32)])
}

fn exec_span_to_value(input: &str, start: usize, end: usize, include_indices: bool) -> Value {
    let mut match_obj = Object::new_array(vec![s_val(&input[start..end])]);
    match_obj
        .properties
        .insert("index".into(), Value::I32(start as i32));
    match_obj.properties.insert("input".into(), s_val(input));
    match_obj
        .properties
        .insert("groups".into(), empty_groups_object());
    if include_indices {
        let mut indices = Object::new_array(vec![range_to_value(start, end)]);
        indices
            .properties
            .insert("groups".into(), empty_groups_object());
        match_obj.properties.insert(
            "indices".into(),
            Value::Object(Arc::new(Mutex::new(indices))),
        );
    }
    Value::Object(Arc::new(Mutex::new(match_obj)))
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
fn compile(pattern: &str, flags: &str) -> Option<Regex> {
    if special_pattern(pattern, flags).is_some() {
        return None;
    }
    if flags.contains('u') && flags.contains('v') {
        return None;
    }
    let normalized_pattern = pattern.replace("(?P<", "(?<");
    let compile_flags: String = flags
        .chars()
        .filter(|c| matches!(c, 'i' | 'm' | 's' | 'u' | 'v'))
        .collect();
    Regex::with_flags(&normalized_pattern, compile_flags.as_str()).ok()
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
    let mut groups = Object::new();
    let mut group_order: Vec<Value> = Vec::new();
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
            Value::Object(Arc::new(Mutex::new(Object::new_array(group_order)))),
        );
    }
    Value::Object(Arc::new(Mutex::new(groups)))
}

fn named_groups_indices_object(m: &Match, index_offset: usize) -> Value {
    let mut groups = Object::new();
    let mut group_order: Vec<Value> = Vec::new();
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
            Value::Object(Arc::new(Mutex::new(Object::new_array(group_order)))),
        );
    }
    Value::Object(Arc::new(Mutex::new(groups)))
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
        indices_obj.properties.insert(
            "groups".into(),
            named_groups_indices_object(m, index_offset),
        );
        match_obj.properties.insert(
            "indices".into(),
            Value::Object(Arc::new(Mutex::new(indices_obj))),
        );
    }
    Value::Object(Arc::new(Mutex::new(match_obj)))
}

fn make_array(elements: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
}

fn s_val(s: &str) -> Value {
    Value::String(Arc::from(s))
}

fn validate_flags(flags: &str) -> Result<(), String> {
    let mut seen = std::collections::HashSet::new();
    for flag in flags.chars() {
        if !matches!(flag, 'd' | 'g' | 'i' | 'm' | 's' | 'u' | 'v' | 'y') {
            return Err(format!("Invalid regular expression flag '{}'", flag));
        }
        if !seen.insert(flag) {
            return Err(format!("Duplicate regular expression flag '{}'", flag));
        }
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
    let normalized_pattern = pattern.replace("(?P<", "(?<");
    let compile_flags: String = flags
        .chars()
        .filter(|c| matches!(c, 'i' | 'm' | 's' | 'u' | 'v'))
        .collect();
    Regex::with_flags(&normalized_pattern, compile_flags.as_str())
        .map(|_| ())
        .map_err(|err| format!("Invalid regular expression: {}", err))
}

fn throw_syntax_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::ecma::error::new_error("SyntaxError", message));
    Value::Null
}

fn throw_type_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::ecma::error::new_error("TypeError", message));
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
                Some(other) => format!("{}", other),
            };
            if let Err(message) = validate_flags(&flags) {
                return throw_syntax_error(ctx, &message);
            }
            if let Err(message) = validate_pattern(&pattern, &flags) {
                return throw_syntax_error(ctx, &message);
            }
            let mut obj = Object::new();
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
                .insert("__type".into(), Value::String(Arc::from(REGEXP_TYPE)));
            obj.properties
                .insert("__proto__".into(), shared_regexp_prototype());
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // newWithFlags(pattern, flags) — explicit flags alias.
    vm.register_host_fn(
        "ecma:regexp",
        "newWithFlags",
        Box::new(|ctx, args| {
            let (pattern, _) = extract_pattern(args, 0);
            let flags = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            if let Err(message) = validate_flags(&flags) {
                return throw_syntax_error(ctx, &message);
            }
            if let Err(message) = validate_pattern(&pattern, &flags) {
                return throw_syntax_error(ctx, &message);
            }
            let mut obj = Object::new();
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
                .insert("__type".into(), Value::String(Arc::from(REGEXP_TYPE)));
            obj.properties
                .insert("__proto__".into(), shared_regexp_prototype());
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // escape(str) — ES2025 §22.2.2.1. Escape all special regex chars.
    vm.register_host_fn(
        "ecma:regexp",
        "escape",
        Box::new(|_ctx, args| {
            let s = match args.first() {
                Some(Value::String(s)) => s.as_ref().to_string(),
                Some(v) => format!("{}", v),
                None => return Value::Undefined,
            };
            let escaped: String = s
                .chars()
                .map(|c| {
                    if c.is_alphanumeric() || c == '_' {
                        c.to_string()
                    } else {
                        format!("\\{}", c)
                    }
                })
                .collect();
            Value::String(Arc::from(escaped.as_str()))
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
        Box::new(|_ctx, args| regexp_test(args)),
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
        Box::new(|_ctx, args| regexp_exec(args)),
    );

    // `regex.toString()` — ECMA-262 §22.2.5.17. Returns "/source/flags".
    vm.register_host_fn(
        "ecma:regexp",
        "toString",
        Box::new(|_ctx, args| regexp_to_string(args)),
    );
}

pub fn dispatch_regexp_method(method: &str, args: &[Value]) -> Option<Value> {
    match method {
        "test" => Some(regexp_test(args)),
        "exec" => Some(regexp_exec(args)),
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

fn regexp_test(args: &[Value]) -> Value {
    Value::Bool(!matches!(regexp_exec(args), Value::Null))
}

fn regexp_exec(args: &[Value]) -> Value {
    let (pattern, flags) = extract_pattern(args, 0);
    let input = s_arg(args, 1);
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
    s_val(&format!("/{}/{}", pattern, flags))
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
    let input = s_arg(args, 0);
    if let Some(method) = args
        .get(1)
        .and_then(|value| lookup_symbol_method(value, "symbolmatch"))
    {
        return ctx.invoke(&method, &[Value::String(Arc::from(input.as_str()))]);
    }
    let (pattern, flags) = extract_pattern(args, 1);
    if let Some(kind) = special_pattern(&pattern, &flags) {
        if flags.contains('g') {
            let matches: Vec<Value> = special_find_all(&input, kind)
                .into_iter()
                .map(|(start, end)| s_val(&input[start..end]))
                .collect();
            return if matches.is_empty() {
                Value::Null
            } else {
                make_array(matches)
            };
        }
        return match special_find(&input, kind, 0) {
            Some((start, end)) => exec_span_to_value(&input, start, end, flags.contains('d')),
            None => Value::Null,
        };
    }
    let re = match compile(&pattern, &flags) {
        Some(re) => re,
        None => return Value::Null,
    };
    if flags.contains('g') {
        let matches: Vec<Value> = re
            .find_iter(&input)
            .map(|m| s_val(m.as_str(&input)))
            .collect();
        if matches.is_empty() {
            Value::Null
        } else {
            make_array(matches)
        }
    } else {
        match re.find(&input) {
            Some(m) => exec_match_to_value(&m, &input, 0, flags.contains('d')),
            None => Value::Null,
        }
    }
}

fn regexp_string_match_all(ctx: &mut HostContext, args: &[Value]) -> Value {
    let input = s_arg(args, 0);
    let (pattern, flags) = extract_pattern(args, 1);
    if regex_like_arg(args.get(1)) && !flags.contains('g') {
        return throw_type_error(
            ctx,
            "String.prototype.matchAll called with a non-global RegExp argument",
        );
    }
    if let Some(kind) = special_pattern(&pattern, &flags) {
        let matches = special_find_all(&input, kind)
            .into_iter()
            .map(|(start, end)| exec_span_to_value(&input, start, end, flags.contains('d')))
            .collect();
        return make_array(matches);
    }
    let re = match compile(&pattern, &flags) {
        Some(re) => re,
        None => return make_array(Vec::new()),
    };
    let mut out = Vec::new();
    for m in re.find_iter(&input) {
        out.push(exec_match_to_value(&m, &input, 0, flags.contains('d')));
    }
    make_array(out)
}

fn regexp_string_search(ctx: &mut HostContext, args: &[Value]) -> Value {
    let input = s_arg(args, 0);
    if let Some(method) = args
        .get(1)
        .and_then(|value| lookup_symbol_method(value, "symbolsearch"))
    {
        return ctx.invoke(&method, &[Value::String(Arc::from(input.as_str()))]);
    }
    let (pattern, flags) = extract_pattern(args, 1);
    if let Some(kind) = special_pattern(&pattern, &flags) {
        return match special_find(&input, kind, 0) {
            Some((start, _)) => Value::I32(start as i32),
            None => Value::I32(-1),
        };
    }
    match compile(&pattern, &flags) {
        Some(re) => match re.find(&input) {
            Some(m) => Value::I32(m.start() as i32),
            None => Value::I32(-1),
        },
        None => Value::I32(-1),
    }
}

fn regexp_string_replace(ctx: &mut HostContext, args: &[Value]) -> Value {
    let input = s_arg(args, 0);
    let (pattern, flags) = extract_pattern(args, 1);
    let re = match compile(&pattern, &flags) {
        Some(re) => re,
        None => return s_val(&input),
    };
    let global = flags.contains('g');
    let replacement_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
    let is_callable = matches!(&replacement_arg, Value::Object(o)
        if matches!(o.lock().unwrap().kind,
            vybe_bytecode::value::ObjectKind::Function(_)
            | vybe_bytecode::value::ObjectKind::HostFunction(_)));
    if is_callable {
        let mut out = String::with_capacity(input.len());
        let mut last_end = 0;
        let matches: Vec<Match> = if global {
            re.find_iter(&input).collect()
        } else {
            re.find(&input).into_iter().collect()
        };
        for m in matches {
            out.push_str(&input[last_end..m.start()]);
            let mut cb_args: Vec<Value> = Vec::with_capacity(m.captures.len() + 3);
            for i in 0..=m.captures.len() {
                cb_args.push(match_group_value(&m, &input, i));
            }
            cb_args.push(Value::I32(m.start() as i32));
            cb_args.push(s_val(&input));
            if m.named_groups().next().is_some() {
                cb_args.push(named_groups_object(&m, &input));
            }
            let ret = ctx.invoke(&replacement_arg, &cb_args);
            match ret {
                Value::String(s) => out.push_str(s.as_ref()),
                other => out.push_str(&format!("{}", other)),
            }
            last_end = m.end();
            if !global {
                break;
            }
        }
        out.push_str(&input[last_end..]);
        return s_val(&out);
    }
    let replacement = s_arg(args, 2);
    s_val(&apply_string_replacement(
        &input,
        &re,
        global,
        replacement.as_str(),
    ))
}

fn regexp_string_replace_all(args: &[Value]) -> Value {
    let input = s_arg(args, 0);
    let (pattern, flags) = extract_pattern(args, 1);
    let replacement = s_arg(args, 2);
    match compile(&pattern, &flags) {
        Some(re) => s_val(&apply_string_replacement(
            &input,
            &re,
            true,
            replacement.as_str(),
        )),
        None => s_val(&input),
    }
}

fn regexp_string_split(args: &[Value]) -> Value {
    let input = s_arg(args, 0);
    let (pattern, flags) = extract_pattern(args, 1);
    let limit = args.get(2).map(|v| v.as_i32().max(0) as usize);
    if matches!(limit, Some(0)) {
        return make_array(Vec::new());
    }
    let max_parts = limit.unwrap_or(usize::MAX);
    if pattern.is_empty() || pattern == "(?:)" {
        return make_array(
            input
                .chars()
                .take(max_parts)
                .map(|ch| s_val(&ch.to_string()))
                .collect(),
        );
    }
    match compile(&pattern, &flags) {
        Some(re) => {
            let mut parts: Vec<Value> = Vec::new();
            let mut last_end = 0;
            for m in re.find_iter(&input) {
                if parts.len() >= max_parts {
                    break;
                }
                parts.push(s_val(&input[last_end..m.start()]));
                for index in 1..=m.captures.len() {
                    if parts.len() >= max_parts {
                        break;
                    }
                    parts.push(match_group_value(&m, &input, index));
                }
                last_end = m.end();
            }
            if parts.len() < max_parts {
                parts.push(s_val(&input[last_end..]));
            }
            make_array(parts)
        }
        None => make_array(vec![s_val(&input)]),
    }
}

fn regex_like_arg(arg: Option<&Value>) -> bool {
    let Some(Value::Object(obj)) = arg else {
        return false;
    };
    let guard = obj.lock().unwrap();
    matches!(guard.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == REGEXP_TYPE)
}

fn apply_string_replacement(input: &str, re: &Regex, global: bool, replacement: &str) -> String {
    let matches: Vec<Match> = if global {
        re.find_iter(input).collect()
    } else {
        re.find(input).into_iter().collect()
    };
    if matches.is_empty() {
        return input.to_string();
    }

    let mut out = String::with_capacity(input.len());
    let mut last_end = 0;
    for m in matches {
        out.push_str(&input[last_end..m.start()]);
        out.push_str(&expand_js_replacement(replacement, input, &m));
        last_end = m.end();
        if !global {
            break;
        }
    }
    out.push_str(&input[last_end..]);
    out
}

fn expand_js_replacement(template: &str, input: &str, m: &Match) -> String {
    let mut out = String::new();
    let chars: Vec<char> = template.chars().collect();
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
    out
}
