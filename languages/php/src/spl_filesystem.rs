//! File-backed SPL iterator operations shared by PHP programs.
use std::fs;
use std::path::PathBuf;
use std::sync::Arc;

use regex::RegexBuilder;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value, heap};

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn("php:spl", "directoryEntries", Box::new(|_, args| {
        let [Value::String(root), flags, ..] = args else { return Value::Bool(false); };
        let Ok(children) = fs::read_dir(root.as_ref()) else { return Value::Bool(false); };
        let flags = flags.as_f64() as u32;
        let mut entries = Vec::new();
        if flags & 4096 == 0 {
            entries.push(PathBuf::from(root.as_ref()).join("."));
            entries.push(PathBuf::from(root.as_ref()).join(".."));
        }
        for child in children {
            let Ok(child) = child else { return Value::Bool(false); };
            entries.push(child.path());
        }
        Value::Object(heap::alloc(Object::new_array(entries.into_iter().map(|path| {
            let path = path.to_string_lossy();
            let path = if cfg!(windows) && flags & 8192 != 0 { path.replace('\\', "/") } else { path.into_owned() };
            Value::String(path.into())
        }).collect())))
    }));
    for (name, validity) in [("directoryEntry", false), ("directoryValid", true)] {
        fw.register_host_fn("php:spl", name, Box::new(move |_, args| {
            let [Value::Object(entries), index, ..] = args else {
                return if validity { Value::Bool(false) } else { Value::String("".into()) };
            };
            let index = index.as_f64();
            let entries = entries.lock().unwrap();
            let ObjectKind::Array(values) = &entries.kind else { return Value::Bool(false); };
            let value = if index.is_finite() && index >= 0.0 && index.fract() == 0.0 {
                values.get(index as usize)
            } else { None };
            if validity { Value::Bool(value.is_some()) }
            else { value.cloned().unwrap_or_else(|| Value::String("".into())) }
        }));
    }
    fw.register_host_fn("php:spl", "isDot", Box::new(|_, args| {
        let Some(Value::String(path)) = args.first() else { return Value::Bool(false); };
        let name = path.rsplit(['/', '\\']).next().unwrap_or("");
        Value::Bool(matches!(name, "." | ".."))
    }));
    fw.register_host_fn("php:spl", "hasChildren", Box::new(|_, args| {
        let [Value::String(path), flags, allow_links, ..] = args else { return Value::Bool(false); };
        let name = path.rsplit(['/', '\\']).next().unwrap_or("");
        if matches!(name, "." | "..") { return Value::Bool(false); }
        let Ok(metadata) = fs::symlink_metadata(path.as_ref()) else { return Value::Bool(false); };
        if metadata.file_type().is_symlink() {
            if flags.as_f64() as u32 & 16384 == 0 && !matches!(allow_links, Value::Bool(true)) {
                return Value::Bool(false);
            }
            return Value::Bool(fs::metadata(path.as_ref()).is_ok_and(|metadata| metadata.is_dir()));
        }
        Value::Bool(metadata.is_dir())
    }));
    fw.register_host_fn("php:spl", "extension", Box::new(|_, args| {
        let Some(Value::String(path)) = args.first() else { return Value::String("".into()); };
        Value::String(std::path::Path::new(path.as_ref()).extension()
            .map(|extension| extension.to_string_lossy().into_owned()).unwrap_or_default().into())
    }));
    fw.register_host_fn("php:spl", "relativePath", Box::new(|_, args| {
        let [Value::String(root), Value::String(path), flags, ..] = args else { return Value::String("".into()); };
        let relative = std::path::Path::new(path.as_ref()).strip_prefix(root.as_ref())
            .map(|path| path.to_string_lossy().into_owned()).unwrap_or_else(|_| path.to_string());
        Value::String(if cfg!(windows) && flags.as_f64() as u32 & 8192 != 0 {
            relative.replace('\\', "/")
        } else { relative }.into())
    }));
    fw.register_host_fn("php:spl", "setTop", Box::new(|_, args| {
        let [Value::Object(values), value, ..] = args else { return Value::Null; };
        let mut values = values.lock().unwrap();
        if let ObjectKind::Array(entries) = &mut values.kind {
            if let Some(last) = entries.last_mut() { *last = value.clone(); }
        }
        Value::Null
    }));
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
