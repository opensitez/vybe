//! `node:path` — Node.js built-in `path` module.
//!
//! Reference: <https://nodejs.org/api/path.html>.
//!
//! Cross-style accessors (`path.posix` / `path.win32`) are deferred;
//! Phase 1 dispatches on the host platform — POSIX on Unix, Windows
//! on Windows.

use std::path::{Component, Path, PathBuf};
use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{VM, Value};

fn s_arg(args: &[Value], idx: usize, default: &str) -> String {
    match args.get(idx) {
        Some(Value::String(text)) => text.to_string(),
        Some(other) => format!("{}", other),
        None => default.to_string() }
}

fn collect_string_args(args: &[Value]) -> Vec<String> {
    args.iter()
        .map(|value| match value {
            Value::String(text) => text.to_string(),
            other => format!("{}", other) })
        .collect()
}

fn s_val(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn sep() -> &'static str {
    if cfg!(windows) { "\\" } else { "/" }
}

fn delimiter() -> &'static str {
    if cfg!(windows) { ";" } else { ":" }
}

/// Node `basename(path[, ext])` — last segment, optionally stripping a
/// trailing extension if it matches.
fn basename(path: &str, ext: Option<&str>) -> String {
    // Trim trailing separators (per Node), then take the last segment.
    let trimmed = path.trim_end_matches(['/', '\\'].as_ref());
    let raw = match trimmed.rsplit_once(|c| c == '/' || c == '\\') {
        Some((_, last)) => last,
        None => trimmed };
    if let Some(e) = ext {
        if !e.is_empty() && raw.ends_with(e) && raw.len() > e.len() {
            return raw[..raw.len() - e.len()].to_string();
        }
    }
    raw.to_string()
}

/// Node `dirname(path)` — parent directory, or `"."` for unqualified
/// names, `"/"` (or `"\\"`) for the filesystem root.
fn dirname(path: &str) -> String {
    let trimmed = path.trim_end_matches(['/', '\\'].as_ref());
    match trimmed.rsplit_once(|c| c == '/' || c == '\\') {
        Some(("", _)) => sep().to_string(),
        Some((dir, _)) => dir.to_string(),
        None => ".".to_string() }
}

/// Node `extname(path)` — `.ext` of the last segment, including the dot.
/// Empty if no dot or if the dot is the first character of the segment
/// (dotfile).
fn extname(path: &str) -> String {
    let base = basename(path, None);
    match base.rfind('.') {
        Some(0) => String::new(), // ".bashrc" → ""
        Some(idx) => base[idx..].to_string(),
        None => String::new() }
}

/// Node `join(...paths)` — concatenate with the platform separator and
/// normalize. Empty segments are skipped (per Node).
fn join(parts: &[String]) -> String {
    let parts: Vec<&str> = parts
        .iter()
        .map(|s| s.as_str())
        .filter(|s| !s.is_empty())
        .collect();
    if parts.is_empty() {
        return ".".to_string();
    }
    let raw = parts.join(sep());
    let normalized = normalize(&raw);
    // Preserve a trailing separator if the original last segment had one
    // (Node's behaviour).
    if parts
        .last()
        .map(|s| s.ends_with('/') || s.ends_with('\\'))
        .unwrap_or(false)
        && !normalized.ends_with('/')
        && !normalized.ends_with('\\')
    {
        format!("{}{}", normalized, sep())
    } else {
        normalized
    }
}

/// Node `normalize(path)` — collapse `..`, `.`, and double separators.
/// Preserves whether the path is absolute and whether it had a trailing
/// separator.
fn normalize(path: &str) -> String {
    if path.is_empty() {
        return ".".to_string();
    }
    let is_abs = path.starts_with('/') || path.starts_with('\\');
    let trailing_sep = path.ends_with('/') || path.ends_with('\\');

    let mut stack: Vec<&str> = Vec::new();
    for seg in path.split(|c| c == '/' || c == '\\') {
        match seg {
            "" | "." => continue,
            ".." => {
                if let Some(top) = stack.last() {
                    if *top != ".." {
                        stack.pop();
                        continue;
                    }
                }
                if !is_abs {
                    stack.push("..");
                }
            }
            other => stack.push(other) }
    }
    let mut joined = stack.join(sep());
    if is_abs {
        joined.insert_str(0, sep());
    }
    if joined.is_empty() {
        return if is_abs {
            sep().to_string()
        } else {
            ".".to_string()
        };
    }
    if trailing_sep && !joined.ends_with('/') && !joined.ends_with('\\') {
        joined.push_str(sep());
    }
    joined
}

/// Node `isAbsolute(path)`.
fn is_absolute(path: &str) -> bool {
    if cfg!(windows) {
        // Drive letter (`C:\foo`) or UNC (`\\srv\share`) — punt on
        // exact Windows rules; cover the common case.
        path.starts_with('\\')
            || path.starts_with('/')
            || (path.len() >= 2 && path.as_bytes()[1] == b':')
    } else {
        path.starts_with('/')
    }
}

/// Node `resolve(...paths)` — fold left, treating each absolute segment
/// as a reset. Empty result resolves to cwd.
fn resolve(parts: &[String]) -> String {
    let mut acc = std::env::current_dir()
        .map(|p| p.to_string_lossy().to_string())
        .unwrap_or_else(|_| ".".to_string());
    for part in parts {
        if part.is_empty() {
            continue;
        }
        if is_absolute(part) {
            acc = part.clone();
        } else {
            acc = format!("{}{}{}", acc, sep(), part);
        }
    }
    let n = normalize(&acc);
    // Strip trailing separator unless it's the root.
    if n.len() > 1 && (n.ends_with('/') || n.ends_with('\\')) {
        n[..n.len() - 1].to_string()
    } else {
        n
    }
}

/// Node `relative(from, to)` — relative path from `from` to `to`.
fn relative(from: &str, to: &str) -> String {
    let from_resolved = PathBuf::from(resolve(&[from.to_string()]));
    let to_resolved = PathBuf::from(resolve(&[to.to_string()]));

    let from_components: Vec<Component> = from_resolved.components().collect();
    let to_components: Vec<Component> = to_resolved.components().collect();

    let common = from_components
        .iter()
        .zip(to_components.iter())
        .take_while(|(a, b)| a == b)
        .count();

    let up_count = from_components.len() - common;
    let mut segments: Vec<String> = Vec::with_capacity(up_count + (to_components.len() - common));
    for _ in 0..up_count {
        segments.push("..".to_string());
    }
    for component in &to_components[common..] {
        segments.push(component.as_os_str().to_string_lossy().to_string());
    }
    segments.join(sep())
}

/// Node `parse(path)` → `{ root, dir, base, ext, name }`.
fn parse(path: &str) -> Value {
    let mut o = Object::new();
    let dir = dirname(path);
    let base = basename(path, None);
    let ext = extname(path);
    let name = if ext.is_empty() {
        base.clone()
    } else {
        base[..base.len() - ext.len()].to_string()
    };
    let root = if path.starts_with('/') {
        "/".to_string()
    } else if cfg!(windows) && path.len() >= 2 && path.as_bytes()[1] == b':' {
        path[..3].to_string()
    } else {
        String::new()
    };
    o.properties.insert("root".into(), s_val(&root));
    o.properties.insert("dir".into(), s_val(&dir));
    o.properties.insert("base".into(), s_val(&base));
    o.properties.insert("ext".into(), s_val(&ext));
    o.properties.insert("name".into(), s_val(&name));
    Value::Object(vybe_runtime::heap::alloc(o))
}

/// Node `format(pathObject)` — opposite of parse.
fn format(obj: &Value) -> String {
    let Value::Object(object) = obj else {
        return String::new();
    };
    let object = object.lock().unwrap();
    let dir = match object.properties.get("dir") {
        Some(Value::String(text)) => text.to_string(),
        _ => match object.properties.get("root") {
            Some(Value::String(text)) => text.to_string(),
            _ => String::new() } };
    let base = match object.properties.get("base") {
        Some(Value::String(text)) => text.to_string(),
        _ => {
            let name = match object.properties.get("name") {
                Some(Value::String(text)) => text.to_string(),
                _ => String::new() };
            let ext = match object.properties.get("ext") {
                Some(Value::String(text)) => text.to_string(),
                _ => String::new() };
            format!("{}{}", name, ext)
        }
    };
    if dir.is_empty() {
        base
    } else if dir.ends_with('/') || dir.ends_with('\\') {
        format!("{}{}", dir, base)
    } else {
        format!("{}{}{}", dir, sep(), base)
    }
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:path",
        "basename",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let ext = match args.get(1) {
                Some(Value::String(text)) => Some(text.to_string()),
                _ => None };
            s_val(&basename(&path, ext.as_deref()))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "dirname",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            s_val(&dirname(&path))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "extname",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            s_val(&extname(&path))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "join",
        Box::new(|_ctx, args| {
            let parts = collect_string_args(args);
            s_val(&join(&parts))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "normalize",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            s_val(&normalize(&path))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "isAbsolute",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            Value::Bool(is_absolute(&path))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "resolve",
        Box::new(|_ctx, args| {
            let parts = collect_string_args(args);
            s_val(&resolve(&parts))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "relative",
        Box::new(|_ctx, args| {
            let from = s_arg(args, 0, "");
            let to = s_arg(args, 1, "");
            s_val(&relative(&from, &to))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "parse",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            parse(&path)
        }),
    );

    vm.register_host_fn(
        "node:path",
        "format",
        Box::new(|_ctx, args| {
            if let Some(obj) = args.first() {
                s_val(&format(obj))
            } else {
                s_val("")
            }
        }),
    );

    vm.register_host_fn("node:path", "sep", Box::new(|_ctx, _args| s_val(sep())));
    vm.register_host_fn(
        "node:path",
        "delimiter",
        Box::new(|_ctx, _args| s_val(delimiter())),
    );

    vm.register_host_fn(
        "node:path",
        "toNamespacedPath",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            if cfg!(windows) && path.starts_with("\\\\") {
                s_val(&format!("\\\\?\\{}", &path[2..]))
            } else {
                s_val(&path)
            }
        }),
    );

    vm.register_host_fn(
        "node:path",
        "posix",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("sep".into(), s_val("/"));
            o.properties.insert("delimiter".into(), s_val(":"));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "win32",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("sep".into(), s_val("\\"));
            o.properties.insert("delimiter".into(), s_val(";"));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:path",
        "matchesGlob",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let pattern = s_arg(args, 1, "");
            Value::Bool(glob_matches(&pattern, &path))
        }),
    );
}

fn glob_matches(pattern: &str, path: &str) -> bool {
    // Convert glob to regex: ** → .*, * → [^/]*, ? → [^/]
    let mut re = String::from("^");
    let mut chars = pattern.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '*' if chars.peek() == Some(&'*') => {
                chars.next();
                re.push_str(".*");
            }
            '*' => re.push_str("[^/]*"),
            '?' => re.push_str("[^/]"),
            '.' | '+' | '(' | ')' | '[' | ']' | '{' | '}' | '^' | '$' | '|' | '\\' => {
                re.push('\\');
                re.push(c);
            }
            other => re.push(other) }
    }
    re.push('$');
    regex::Regex::new(&re)
        .map(|r| r.is_match(path))
        .unwrap_or(false)
}

#[allow(dead_code)]
fn _force_use(_: &Path, _: ObjectKind) {}
