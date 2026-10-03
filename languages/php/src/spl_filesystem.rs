//! File-backed SPL iterator operations shared by PHP programs.
use std::fs;
use std::path::PathBuf;
use std::sync::Arc;

use regex::RegexBuilder;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value, heap};

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:filesystem",
        "glob",
        Box::new(|_, args| {
            let Some(Value::String(pattern)) = args.first() else {
                return Value::Bool(false);
            };
            // PHP/POSIX glob treats repeated stars within a path component as
            // one wildcard, not the recursive globstar extension.
            let mut normalized = String::new();
            let mut chars = pattern.chars().peekable();
            while let Some(ch) = chars.next() {
                if ch == '\\' {
                    if let Some(literal) = chars.next() {
                        normalized.push_str(&glob::Pattern::escape(&literal.to_string()));
                    } else {
                        normalized.push(ch);
                    }
                } else {
                    normalized.push(ch);
                    if ch == '*' {
                        while chars.peek() == Some(&'*') {
                            chars.next();
                        }
                    }
                }
            }
            let options = glob::MatchOptions {
                case_sensitive: true,
                require_literal_separator: true,
                require_literal_leading_dot: true,
            };
            let Ok(paths) = glob::glob_with(&normalized, options) else {
                return Value::Bool(false);
            };
            let mut paths: Vec<String> = paths
                .filter_map(Result::ok)
                .map(|path| path.to_string_lossy().into_owned())
                .collect();
            paths.sort();
            Value::Object(heap::alloc(Object::new_array(
                paths
                    .into_iter()
                    .map(|path| Value::String(Arc::from(path)))
                    .collect(),
            )))
        }),
    );
    fw.register_host_fn(
        "php:spl",
        "recursiveFiles",
        Box::new(|_, args| {
            let Some(Value::String(root)) = args.first() else {
                return Value::Null;
            };
            let mut result = Object::new();
            result.kind = ObjectKind::Map(Default::default());
            let mut pending = vec![PathBuf::from(root.as_ref())];
            while let Some(directory) = pending.pop() {
                let Ok(children) = fs::read_dir(directory) else {
                    continue;
                };
                for child in children.flatten() {
                    let path = child.path();
                    let Ok(kind) = child.file_type() else {
                        continue;
                    };
                    if kind.is_dir() {
                        pending.push(path);
                    } else if kind.is_file() {
                        let path: Arc<str> = Arc::from(path.to_string_lossy().into_owned());
                        if let ObjectKind::Map(entries) = &mut result.kind {
                            entries.insert(Value::String(path.clone()), Value::String(path));
                        }
                    }
                }
            }
            Value::Object(heap::alloc(result))
        }),
    );

    fw.register_host_fn(
        "php:spl",
        "regexFilter",
        Box::new(|_, args| {
            let [Value::Object(source), Value::String(pattern), ..] = args else {
                return Value::Null;
            };
            let text = pattern.as_ref();
            let Some(delimiter) = text.chars().next() else {
                return Value::Null;
            };
            let Some(end) = text.rfind(delimiter).filter(|end| *end > 0) else {
                return Value::Null;
            };
            let flags = &text[end + delimiter.len_utf8()..];
            let Ok(regex) = RegexBuilder::new(&text[delimiter.len_utf8()..end])
                .case_insensitive(flags.contains('i'))
                .multi_line(flags.contains('m'))
                .dot_matches_new_line(flags.contains('s'))
                .build()
            else {
                return Value::Null;
            };
            let values: Vec<(Value, Value)> = {
                let object = source.lock().unwrap();
                match &object.kind {
                    ObjectKind::Map(entries) => entries
                        .iter()
                        .map(|(k, v)| (k.clone(), v.clone()))
                        .collect(),
                    ObjectKind::Array(items) => items
                        .iter()
                        .enumerate()
                        .map(|(i, v)| (Value::I64(i as i64), v.clone()))
                        .collect(),
                    _ => return Value::Null,
                }
            };
            let mut result = Object::new();
            result.kind = ObjectKind::Map(Default::default());
            if let ObjectKind::Map(entries) = &mut result.kind {
                for (key, value) in values {
                    if let Value::String(text) = &value {
                        if regex.is_match(text) {
                            entries.insert(key, value);
                        }
                    }
                }
            }
            Value::Object(heap::alloc(result))
        }),
    );
}
