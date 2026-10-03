//! PHP's ordered pattern/replacement arrays, before scalar regex execution.
use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value, heap};

struct Pattern {
    regex: pcre2::bytes::Regex,
    nonempty_regex: Option<pcre2::bytes::Regex>,
    anchored: bool,
    utf: bool,
}

fn compile_pattern(literal: &str) -> Option<Pattern> {
    let literal = literal.trim_start();
    let bytes = literal.as_bytes();
    let &open = bytes.first()?;
    if !open.is_ascii()
        || open.is_ascii_alphanumeric()
        || open.is_ascii_whitespace()
        || open == b'\\'
    {
        return None;
    }
    let close = match open {
        b'(' => b')',
        b'[' => b']',
        b'{' => b'}',
        b'<' => b'>',
        _ => open,
    };
    let mut depth = 1;
    let mut cursor = 1;
    while cursor < bytes.len() {
        if bytes[cursor] == b'\\' {
            cursor += 2;
            continue;
        }
        if open != close && bytes[cursor] == open {
            depth += 1;
        }
        if bytes[cursor] == close {
            depth -= 1;
            if depth == 0 {
                break;
            }
        }
        cursor += 1;
    }
    if cursor >= bytes.len() {
        return None;
    }
    let flags = literal[cursor + 1..].trim();
    let mut builder = pcre2::bytes::RegexBuilder::new();
    builder
        .caseless(flags.contains('i'))
        .multi_line(flags.contains('m'))
        .dotall(flags.contains('s'))
        .extended(flags.contains('x'))
        .utf(flags.contains('u'))
        .ucp(flags.contains('u'))
        .jit_if_available(true);
    let mut body = String::new();
    for flag in flags.chars() {
        match flag {
            'U' => body.push_str("(?U)"),
            'J' => body.push_str("(?J)"),
            'n' => body.push_str("(?n)"),
            'i' | 'm' | 's' | 'x' | 'u' | 'A' | 'S' | 'X' => {}
            _ => return None,
        }
    }
    body.push_str(&literal[1..cursor]);
    let regex = builder.build(&body).ok()?;
    // Retry a nonempty alternative at the same search position after an empty
    // match, as PHP does. \G anchors both ends to that search position.
    let suffix = if flags.contains('x') { "\n" } else { "" };
    let nonempty_regex = builder.build(&format!(r"\G(?:{body}{suffix})(?!\G)")).ok();
    Some(Pattern {
        regex,
        nonempty_regex,
        anchored: flags.contains('A'),
        utf: flags.contains('u'),
    })
}

// PHP replacements recognize numeric backreferences, not JavaScript's $& or
// named-group replacement syntax. Work in bytes to preserve capture offsets.
fn expand_replacement(
    output: &mut Vec<u8>,
    replacement: &[u8],
    captures: &pcre2::bytes::CaptureLocations,
    input: &[u8],
) {
    let mut cursor = 0;
    while cursor < replacement.len() {
        let byte = replacement[cursor];
        if byte == b'\\' && matches!(replacement.get(cursor + 1), Some(b'\\' | b'$')) {
            output.push(replacement[cursor + 1]);
            cursor += 2;
            continue;
        }
        if byte == b'$' || byte == b'\\' {
            let mut end = cursor + 1;
            let braced = byte == b'$' && replacement.get(end) == Some(&b'{');
            if braced {
                end += 1;
            }
            let start = end;
            let mut index = 0;
            while end < replacement.len() && end - start < 2 && replacement[end].is_ascii_digit() {
                index = index * 10 + (replacement[end] - b'0') as usize;
                end += 1;
            }
            if end > start && (!braced || replacement.get(end) == Some(&b'}')) {
                if let Some((start, end)) = captures.get(index) {
                    output.extend_from_slice(&input[start..end]);
                }
                cursor = end + usize::from(braced);
                continue;
            }
        }
        output.push(byte);
        cursor += 1;
    }
}

fn replace_scalar(pattern: &Pattern, replacement: &str, input: &str) -> Value {
    let mut output = Vec::with_capacity(input.len());
    let (mut end, mut search, mut retry_nonempty) = (0, 0, false);
    let mut captures = pattern.regex.capture_locations();
    while search <= input.len() {
        let regex = if retry_nonempty {
            pattern.nonempty_regex.as_ref()
        } else {
            Some(&pattern.regex)
        };
        let found = match regex {
            Some(regex) => match regex.captures_read_at(&mut captures, input.as_bytes(), search) {
                Ok(found) => found,
                Err(_) => return Value::Null,
            },
            None => None,
        };
        let Some(found) = found else {
            if !retry_nonempty {
                break;
            }
            search += if pattern.utf {
                input[search..].chars().next().map_or(1, char::len_utf8)
            } else {
                1
            };
            retry_nonempty = false;
            continue;
        };
        if pattern.anchored && found.start() != search {
            break;
        }
        output.extend_from_slice(&input.as_bytes()[end..found.start()]);
        expand_replacement(
            &mut output,
            replacement.as_bytes(),
            &captures,
            input.as_bytes(),
        );
        end = found.end();
        search = end;
        retry_nonempty = found.start() == found.end();
    }
    output.extend_from_slice(&input.as_bytes()[end..]);
    Value::String(Arc::from(String::from_utf8_lossy(&output).as_ref()))
}

fn replace_subject(pattern: &Pattern, replacement: &str, subject: &Value) -> Value {
    if let Value::Object(object) = subject {
        let mut snapshot = object.lock().unwrap().clone();
        match &mut snapshot.kind {
            ObjectKind::Array(items) => {
                for item in items {
                    *item = replace_subject(pattern, replacement, item);
                }
            }
            ObjectKind::Map(items) => {
                for item in items.values_mut() {
                    *item = replace_subject(pattern, replacement, item);
                }
            }
            _ => return replace_scalar(pattern, replacement, &subject.to_string()),
        }
        return Value::Object(heap::alloc(snapshot));
    }
    replace_scalar(pattern, replacement, &subject.to_string())
}

fn match_result(status: Value, matches: Object) -> Value {
    array(vec![status, Value::Object(heap::alloc(matches))])
}

fn match_with_offsets(pattern: &Pattern, input: &str, flags: i32, offset: i64) -> Value {
    let mut matches = Object::new();
    matches.kind = ObjectKind::Map(Default::default());
    let offset = if offset < 0 {
        (input.len() as i64 + offset).max(0)
    } else {
        offset
    };
    if offset > input.len() as i64 {
        return match_result(Value::Bool(false), matches);
    }
    let mut locations = pattern.regex.capture_locations();
    match pattern
        .regex
        .captures_read_at(&mut locations, input.as_bytes(), offset as usize)
    {
        Ok(Some(found)) if !pattern.anchored || found.start() == offset as usize => {}
        Ok(_) => return match_result(Value::I32(0), matches),
        Err(_) => return match_result(Value::Bool(false), matches),
    }
    let unmatched_null = flags & 512 != 0;
    let count = if unmatched_null {
        locations.len()
    } else {
        (0..locations.len())
            .rfind(|&index| locations.get(index).is_some())
            .map_or(0, |index| index + 1)
    };
    let ObjectKind::Map(entries) = &mut matches.kind else {
        unreachable!()
    };
    for index in 0..count {
        let range = locations.get(index);
        let text = match range {
            Some((start, end)) => Value::String(Arc::from(
                String::from_utf8_lossy(&input.as_bytes()[start..end]).as_ref(),
            )),
            None if unmatched_null => Value::Null,
            None => Value::String(Arc::from("")),
        };
        let value = if flags & 256 != 0 {
            array(vec![
                text,
                Value::I64(range.map_or(-1, |(start, _)| start as i64)),
            ])
        } else {
            text
        };
        if let Some(name) = &pattern.regex.capture_names()[index] {
            let key = Value::String(Arc::from(name.as_str()));
            if range.is_some() {
                entries.insert(key, value.clone());
            } else {
                entries.entry(key).or_insert_with(|| value.clone());
            }
        }
        entries.insert(Value::I32(index as i32), value);
    }
    match_result(Value::I32(1), matches)
}

// PCRE matching for both PHP preg_match_all layouts. In particular, lookaround
// and Unicode classes must use the same engine as the other PHP regex APIs.
fn match_all_groups(
    pattern: &Pattern,
    input: &str,
    set_order: bool,
    offset_capture: bool,
) -> Value {
    let mut columns = vec![Vec::new(); pattern.regex.captures_len()];
    let mut rows = Vec::new();
    let mut captures = pattern.regex.capture_locations();
    let (mut search, mut retry_nonempty) = (0, false);
    while search <= input.len() {
        let regex = if retry_nonempty {
            pattern.nonempty_regex.as_ref()
        } else {
            Some(&pattern.regex)
        };
        let found = match regex {
            Some(regex) => match regex.captures_read_at(&mut captures, input.as_bytes(), search) {
                Ok(found) => found,
                Err(_) => return Value::Bool(false),
            },
            None => None,
        };
        let Some(found) = found else {
            if !retry_nonempty {
                break;
            }
            search += if pattern.utf {
                input[search..].chars().next().map_or(1, char::len_utf8)
            } else {
                1
            };
            retry_nonempty = false;
            continue;
        };
        if pattern.anchored && found.start() != search {
            break;
        }
        let count = if set_order {
            (0..captures.len())
                .rfind(|&i| captures.get(i).is_some())
                .map_or(0, |i| i + 1)
        } else {
            captures.len()
        };
        let mut row = Object::new();
        row.kind = ObjectKind::Map(Default::default());
        let ObjectKind::Map(entries) = &mut row.kind else {
            unreachable!()
        };
        for i in 0..count {
            let capture = captures.get(i);
            let text = capture.map_or_else(
                || Value::String(Arc::from("")),
                |(start, end)| {
                    Value::String(Arc::from(
                        String::from_utf8_lossy(&input.as_bytes()[start..end]).as_ref(),
                    ))
                },
            );
            let text = if offset_capture {
                array(vec![
                    text,
                    Value::I64(capture.map_or(-1, |(start, _)| start as i64)),
                ])
            } else {
                text
            };
            if set_order {
                if let Some(name) = &pattern.regex.capture_names()[i] {
                    entries.insert(Value::String(Arc::from(name.as_str())), text.clone());
                }
                entries.insert(Value::I32(i as i32), text);
            } else {
                columns[i].push(text);
            }
        }
        if set_order {
            rows.push(Value::Object(heap::alloc(row)));
        }
        search = found.end();
        retry_nonempty = found.start() == found.end();
    }
    if set_order {
        return array(rows);
    }
    let mut result = Object::new();
    result.kind = ObjectKind::Map(Default::default());
    let ObjectKind::Map(entries) = &mut result.kind else {
        unreachable!()
    };
    for (i, column) in columns.into_iter().enumerate() {
        let column = array(column);
        if let Some(name) = &pattern.regex.capture_names()[i] {
            entries.insert(Value::String(Arc::from(name.as_str())), column.clone());
        }
        entries.insert(Value::I32(i as i32), column);
    }
    Value::Object(heap::alloc(result))
}

fn split(pattern: &Pattern, input: &str, limit: i64, flags: i32) -> Value {
    let mut result = Vec::new();
    let mut remaining = if limit > 0 {
        limit as usize
    } else {
        usize::MAX
    };
    let mut end = 0;
    let part = |start: usize, end: usize| {
        let text = Value::String(Arc::from(
            String::from_utf8_lossy(&input.as_bytes()[start..end]).as_ref(),
        ));
        if flags & 4 != 0 {
            array(vec![text, Value::I64(start as i64)])
        } else {
            text
        }
    };
    let mut offset = 0;
    let mut previous_end = None;
    let mut capture = pattern.regex.capture_locations();
    while offset <= input.len() && remaining > 1 {
        let matched = match pattern
            .regex
            .captures_read_at(&mut capture, input.as_bytes(), offset)
        {
            Ok(Some(matched)) => matched,
            Ok(None) => break,
            Err(_) => return Value::Bool(false),
        };
        if pattern.anchored && matched.start() != offset {
            break;
        }
        if matched.start() == matched.end() {
            // The byte iterator advances one byte after an empty match. In
            // UTF mode the next search must start at a code point boundary.
            offset = matched.end()
                + if pattern.utf {
                    input[matched.end()..]
                        .chars()
                        .next()
                        .map_or(1, char::len_utf8)
                } else {
                    1
                };
            if previous_end == Some(matched.end()) {
                continue;
            }
        } else {
            offset = matched.end();
        }
        previous_end = Some(matched.end());
        if flags & 1 == 0 || matched.start() > end {
            result.push(part(end, matched.start()));
            remaining -= 1;
        }
        if flags & 2 != 0 {
            let last = (1..capture.len())
                .rfind(|&index| capture.get(index).is_some())
                .unwrap_or(0);
            for index in 1..=last {
                match capture.get(index) {
                    Some((start, end)) if flags & 1 == 0 || start != end => {
                        result.push(part(start, end))
                    }
                    None if flags & 1 == 0 => result.push(if flags & 4 != 0 {
                        array(vec![Value::String(Arc::from("")), Value::I64(-1)])
                    } else {
                        Value::String(Arc::from(""))
                    }),
                    _ => {}
                }
            }
        }
        end = matched.end();
    }
    if flags & 1 == 0 || end < input.len() {
        result.push(part(end, input.len()));
    }
    array(result)
}

type PatternCache = Mutex<HashMap<String, Option<Arc<Pattern>>>>;

fn cached_pattern(cache: &PatternCache, literal: String) -> Option<Arc<Pattern>> {
    let mut cache = cache.lock().unwrap();
    if let Some(pattern) = cache.get(&literal) {
        return pattern.clone();
    }
    let pattern = compile_pattern(&literal).map(Arc::new);
    if cache.len() >= 256 {
        cache.clear();
    }
    cache.insert(literal, pattern.clone());
    pattern
}

fn values(value: &Value) -> Option<Vec<Value>> {
    let Value::Object(object) = value else {
        return None;
    };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(items) => Some(items.clone()),
        ObjectKind::Map(items) => Some(items.values().cloned().collect()),
        _ => None,
    }
}

fn array(items: Vec<Value>) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::Array(items);
    Value::Object(heap::alloc(object))
}

pub fn register(fw: &mut Framework<'_>) {
    let patterns = Arc::new(PatternCache::default());
    let replace_patterns = patterns.clone();
    fw.register_host_fn(
        "php:regex",
        "replaceAll",
        Box::new(move |_, args| {
            let [subject, literal, replacement, ..] = args else {
                return Value::Null;
            };
            match cached_pattern(&replace_patterns, literal.to_string()) {
                Some(pattern) => replace_subject(&pattern, &replacement.to_string(), subject),
                None => Value::Null,
            }
        }),
    );
    for (name, set_order) in [("matchAllGroups", false), ("matchAllSetOrder", true)] {
        let group_patterns = patterns.clone();
        fw.register_host_fn(
            "php:regex",
            name,
            Box::new(move |_, args| {
                let literal = args.first().map(ToString::to_string).unwrap_or_default();
                let input = args.get(1).map(ToString::to_string).unwrap_or_default();
                let flags = match args.get(2) {
                    Some(Value::I32(value)) => *value as i64,
                    Some(Value::I64(value)) => *value,
                    Some(Value::F64(value)) => *value as i64,
                    _ => 0,
                };
                match cached_pattern(&group_patterns, literal) {
                    Some(pattern) => {
                        match_all_groups(&pattern, &input, set_order, flags & 256 != 0)
                    }
                    None => Value::Bool(false),
                }
            }),
        );
    }
    let match_patterns = patterns.clone();
    fw.register_host_fn(
        "php:regex",
        "matchWithOffsets",
        Box::new(move |_, args| {
            let literal = args.first().map(ToString::to_string).unwrap_or_default();
            let input = args.get(1).map(ToString::to_string).unwrap_or_default();
            let flags = args.get(2).map_or(0, Value::as_i32);
            let offset = args.get(3).map_or(0, |value| value.as_f64() as i64);
            let pattern = cached_pattern(&match_patterns, literal);
            match pattern {
                Some(pattern) => match_with_offsets(&pattern, &input, flags, offset),
                None => {
                    let mut empty = Object::new();
                    empty.kind = ObjectKind::Map(Default::default());
                    match_result(Value::Bool(false), empty)
                }
            }
        }),
    );
    fw.register_host_fn(
        "php:regex",
        "split",
        Box::new(move |_, args| {
            let literal = args.first().map(ToString::to_string).unwrap_or_default();
            let input = args.get(1).map(ToString::to_string).unwrap_or_default();
            let limit = args.get(2).map_or(-1, |value| value.as_f64() as i64);
            let flags = args.get(3).map_or(0, Value::as_i32);
            match cached_pattern(&patterns, literal) {
                Some(pattern) => split(&pattern, &input, limit, flags),
                None => Value::Bool(false),
            }
        }),
    );
    fw.register_host_fn(
        "php:regex",
        "replacementPairs",
        Box::new(|_, args| {
            let [patterns, replacement, ..] = args else {
                return Value::Null;
            };
            let Some(patterns) = values(patterns) else {
                return array(vec![array(vec![patterns.clone(), replacement.clone()])]);
            };
            let replacements = values(replacement);
            array(
                patterns
                    .into_iter()
                    .enumerate()
                    .map(|(index, pattern)| {
                        let replacement = match &replacements {
                            Some(items) => items
                                .get(index)
                                .cloned()
                                .unwrap_or_else(|| Value::String("".into())),
                            None => replacement.clone(),
                        };
                        array(vec![pattern, replacement])
                    })
                    .collect(),
            )
        }),
    );
}
