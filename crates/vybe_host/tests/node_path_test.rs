//! Behaviour tests for `node:path` host imports.
//!
//! Reference: <https://nodejs.org/api/path.html>.
//!
//! Notes:
//! - Vybe's host is single-platform per process; tests assume the
//!   POSIX path style on Unix and Windows style on Windows. Tests
//!   that depend on platform-specific output gate via `cfg(unix)` /
//!   `cfg(windows)`.
//! - `path.posix` / `path.win32` cross-style accessors are deferred
//!   to a later phase.

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_path(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-path-test>");
    let import_idx = chunk.add_import("node:path", name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

fn prop(value: &Value, key: &str) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(v) = object.properties.get(key) {
            return v.clone();
        }
    }
    Value::Null
}

// ── basename ──────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn basename_strips_directory() {
    assert_eq!(as_string(&call_path("basename", vec![s("/foo/bar/baz.txt")])), "baz.txt");
}

#[cfg(unix)]
#[test]
fn basename_with_extension_to_strip() {
    assert_eq!(
        as_string(&call_path("basename", vec![s("/foo/bar/baz.txt"), s(".txt")])),
        "baz"
    );
}

#[cfg(unix)]
#[test]
fn basename_returns_string_for_non_matching_ext() {
    // Node: returns the basename unchanged when the ext doesn't match.
    assert_eq!(
        as_string(&call_path("basename", vec![s("/foo/bar/baz.txt"), s(".md")])),
        "baz.txt"
    );
}

#[cfg(unix)]
#[test]
fn basename_handles_trailing_slash() {
    assert_eq!(as_string(&call_path("basename", vec![s("/foo/bar/")])), "bar");
}

// ── dirname ───────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn dirname_returns_parent() {
    assert_eq!(as_string(&call_path("dirname", vec![s("/foo/bar/baz.txt")])), "/foo/bar");
}

#[cfg(unix)]
#[test]
fn dirname_root_returns_root() {
    assert_eq!(as_string(&call_path("dirname", vec![s("/foo")])), "/");
}

#[cfg(unix)]
#[test]
fn dirname_no_separator_returns_dot() {
    assert_eq!(as_string(&call_path("dirname", vec![s("foo")])), ".");
}

// ── extname ───────────────────────────────────────────────────────

#[test]
fn extname_returns_dot_extension() {
    assert_eq!(as_string(&call_path("extname", vec![s("foo.txt")])), ".txt");
}

#[test]
fn extname_no_extension_returns_empty() {
    assert_eq!(as_string(&call_path("extname", vec![s("foo")])), "");
}

#[test]
fn extname_dotfile_no_extension() {
    // Node: `.bashrc` has no extension (no dot before the last segment).
    assert_eq!(as_string(&call_path("extname", vec![s(".bashrc")])), "");
}

#[test]
fn extname_only_returns_last_dot() {
    assert_eq!(as_string(&call_path("extname", vec![s("a.b.c")])), ".c");
}

// ── join ──────────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn join_concatenates_with_separator() {
    assert_eq!(
        as_string(&call_path("join", vec![s("foo"), s("bar"), s("baz.txt")])),
        "foo/bar/baz.txt"
    );
}

#[cfg(unix)]
#[test]
fn join_normalizes_double_slashes() {
    assert_eq!(
        as_string(&call_path("join", vec![s("/foo/"), s("/bar/")])),
        "/foo/bar/"
    );
}

#[cfg(unix)]
#[test]
fn join_handles_dotdot() {
    assert_eq!(
        as_string(&call_path("join", vec![s("/foo/bar"), s(".."), s("baz")])),
        "/foo/baz"
    );
}

// ── normalize ─────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn normalize_collapses_double_slashes() {
    assert_eq!(as_string(&call_path("normalize", vec![s("/foo//bar")])), "/foo/bar");
}

#[cfg(unix)]
#[test]
fn normalize_resolves_dotdot() {
    assert_eq!(as_string(&call_path("normalize", vec![s("/foo/bar/../baz")])), "/foo/baz");
}

#[cfg(unix)]
#[test]
fn normalize_drops_leading_dot_slash() {
    assert_eq!(as_string(&call_path("normalize", vec![s("./foo")])), "foo");
}

// ── isAbsolute ────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn is_absolute_true_for_root_relative() {
    assert_eq!(call_path("isAbsolute", vec![s("/foo")]), Value::Bool(true));
}

#[cfg(unix)]
#[test]
fn is_absolute_false_for_relative() {
    assert_eq!(call_path("isAbsolute", vec![s("foo/bar")]), Value::Bool(false));
}

// ── resolve ───────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn resolve_absolute_passes_through() {
    assert_eq!(as_string(&call_path("resolve", vec![s("/etc/passwd")])), "/etc/passwd");
}

#[cfg(unix)]
#[test]
fn resolve_joins_relative_to_absolute() {
    assert_eq!(
        as_string(&call_path("resolve", vec![s("/foo"), s("bar"), s("baz")])),
        "/foo/bar/baz"
    );
}

#[cfg(unix)]
#[test]
fn resolve_later_absolute_wins() {
    // Node: `path.resolve('/foo', '/bar')` → '/bar' (later abs resets).
    assert_eq!(
        as_string(&call_path("resolve", vec![s("/foo"), s("/bar"), s("baz")])),
        "/bar/baz"
    );
}

// ── relative ──────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn relative_between_siblings() {
    assert_eq!(
        as_string(&call_path("relative", vec![s("/foo/bar"), s("/foo/baz")])),
        "../baz"
    );
}

#[cfg(unix)]
#[test]
fn relative_to_descendant() {
    assert_eq!(
        as_string(&call_path("relative", vec![s("/foo"), s("/foo/bar")])),
        "bar"
    );
}

#[cfg(unix)]
#[test]
fn relative_same_path_empty() {
    assert_eq!(
        as_string(&call_path("relative", vec![s("/foo"), s("/foo")])),
        ""
    );
}

// ── parse ─────────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn parse_returns_components() {
    let parsed = call_path("parse", vec![s("/home/user/file.txt")]);
    assert_eq!(as_string(&prop(&parsed, "root")), "/");
    assert_eq!(as_string(&prop(&parsed, "dir")), "/home/user");
    assert_eq!(as_string(&prop(&parsed, "base")), "file.txt");
    assert_eq!(as_string(&prop(&parsed, "name")), "file");
    assert_eq!(as_string(&prop(&parsed, "ext")), ".txt");
}

// ── format ────────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn format_round_trips_with_parse() {
    let parsed = call_path("parse", vec![s("/home/user/file.txt")]);
    let formatted = as_string(&call_path("format", vec![parsed]));
    assert_eq!(formatted, "/home/user/file.txt");
}

// ── sep / delimiter ───────────────────────────────────────────────

#[test]
fn sep_returns_platform_separator() {
    let v = as_string(&call_path("sep", vec![]));
    if cfg!(windows) {
        assert_eq!(v, "\\");
    } else {
        assert_eq!(v, "/");
    }
}

#[test]
fn delimiter_returns_platform_path_delimiter() {
    let v = as_string(&call_path("delimiter", vec![]));
    if cfg!(windows) {
        assert_eq!(v, ";");
    } else {
        assert_eq!(v, ":");
    }
}

#[allow(dead_code)]
fn _force_object_use(_: Object, _: ObjectKind) {}
