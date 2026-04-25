//! Behaviour tests for `node:process` host imports.
//!
//! Reference: <https://nodejs.org/api/process.html>.
//!
//! Phase 1 covers always-available reads + a few mutators (`cwd`,
//! `chdir`, `exit`-test-deferred). Streams (`stdin`/`stdout`/`stderr`),
//! event listeners (`on('exit', ...)`), and `nextTick` come later.

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_proc(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-process-test>");
    let import_idx = chunk.add_import("node:process", name);
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

fn array_strings(value: &Value) -> Vec<String> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(elements) = &object.kind {
            return elements
                .iter()
                .map(|element| match element {
                    Value::String(text) => text.to_string(),
                    other => format!("{}", other),
                })
                .collect();
        }
    }
    Vec::new()
}

// ── platform / arch / version ─────────────────────────────────────

#[test]
fn platform_returns_recognized_string() {
    let s = as_string(&call_proc("platform", vec![]));
    assert!(
        matches!(s.as_str(), "darwin" | "linux" | "win32" | "freebsd" | "openbsd"),
        "process.platform() should be Node-recognized, got {}",
        s
    );
}

#[test]
fn arch_returns_recognized_string() {
    let s = as_string(&call_proc("arch", vec![]));
    assert!(
        matches!(s.as_str(), "x64" | "arm64" | "ia32" | "arm" | "ppc64" | "s390x" | "riscv64"),
        "process.arch() should be Node-recognized, got {}",
        s
    );
}

#[test]
fn version_returns_v_prefixed_string() {
    // Node uses "vMAJOR.MINOR.PATCH" — Vybe surfaces something with a
    // leading 'v' so consuming code's regex still matches.
    let v = as_string(&call_proc("version", vec![]));
    assert!(v.starts_with('v'), "process.version should start with 'v', got {}", v);
}

// ── pid / ppid ────────────────────────────────────────────────────

#[test]
fn pid_returns_positive_integer() {
    let v = call_proc("pid", vec![]);
    if let Value::F64(n) = v {
        assert!(n > 0.0 && n.is_finite(), "process.pid should be positive int, got {}", n);
        assert_eq!(n.fract(), 0.0, "process.pid should be integer-valued, got {}", n);
    } else {
        panic!("process.pid expected number, got {:?}", v);
    }
}

// ── cwd / chdir ───────────────────────────────────────────────────

#[test]
fn cwd_returns_existing_directory() {
    let s = as_string(&call_proc("cwd", vec![]));
    assert!(
        std::path::Path::new(&s).is_dir(),
        "process.cwd() should point to existing directory, got {}",
        s
    );
}

#[test]
fn chdir_changes_working_directory() {
    // Each test gets a temp dir; chdir into it, verify cwd reflects it,
    // then chdir back so we don't pollute.
    let original = std::env::current_dir().expect("get cwd");
    let dir = std::env::temp_dir().join(format!(
        "vybe-node-process-chdir-{}",
        std::process::id()
    ));
    let _ = std::fs::create_dir_all(&dir);

    call_proc("chdir", vec![s(dir.to_str().unwrap())]);
    let after = as_string(&call_proc("cwd", vec![]));
    assert!(
        std::path::Path::new(&after).canonicalize().ok()
            == dir.canonicalize().ok(),
        "after chdir, cwd should be the new dir; got {} vs {}",
        after,
        dir.display()
    );

    let _ = std::env::set_current_dir(&original);
    let _ = std::fs::remove_dir_all(&dir);
}

// ── argv ──────────────────────────────────────────────────────────

#[test]
fn argv_returns_array_of_strings() {
    let v = call_proc("argv", vec![]);
    let names = array_strings(&v);
    assert!(!names.is_empty(), "process.argv should have at least the executable path");
}

// ── env ───────────────────────────────────────────────────────────

#[test]
fn env_returns_object_with_path_or_userprofile() {
    let v = call_proc("env", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        // Almost every shell has at least PATH (Unix) or Path (Windows).
        let has_path = o.properties.keys().any(|k| k.eq_ignore_ascii_case("path"));
        assert!(has_path, "process.env should have a PATH-like entry");
        return;
    }
    panic!("process.env expected object, got {:?}", v);
}

// ── exec_path / argv0 ─────────────────────────────────────────────

#[test]
fn exec_path_returns_existing_file() {
    let s = as_string(&call_proc("execPath", vec![]));
    assert!(
        std::path::Path::new(&s).exists(),
        "process.execPath should point to an existing file, got {}",
        s
    );
}

// ── uptime / hrtime ───────────────────────────────────────────────

#[test]
fn uptime_returns_non_negative_number() {
    let v = call_proc("uptime", vec![]);
    if let Value::F64(n) = v {
        assert!(n >= 0.0 && n.is_finite(), "uptime should be non-negative finite, got {}", n);
    } else {
        panic!("process.uptime() expected number, got {:?}", v);
    }
}

#[test]
fn hrtime_returns_two_element_array() {
    // Node: hrtime() returns [seconds, nanoseconds] tuple.
    let v = call_proc("hrtime", vec![]);
    if let Value::Object(object) = &v {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(elems) = &object.kind {
            assert_eq!(elems.len(), 2, "hrtime() returns [s, ns] pair");
            assert!(matches!(&elems[0], Value::F64(_) | Value::I32(_)), "hrtime[0] number");
            assert!(matches!(&elems[1], Value::F64(_) | Value::I32(_)), "hrtime[1] number");
            return;
        }
    }
    panic!("process.hrtime() expected 2-array, got {:?}", v);
}

// ── memoryUsage ───────────────────────────────────────────────────

#[test]
fn memory_usage_returns_object_with_rss() {
    let v = call_proc("memoryUsage", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        // Node returns {rss, heapTotal, heapUsed, external, arrayBuffers}.
        // Vybe doesn't have a JS heap so the values are best-effort,
        // but the keys must exist for compat.
        for key in ["rss", "heapTotal", "heapUsed", "external", "arrayBuffers"] {
            assert!(
                o.properties.contains_key(key),
                "process.memoryUsage().{} should exist",
                key
            );
        }
        return;
    }
    panic!("process.memoryUsage() expected object, got {:?}", v);
}

#[allow(dead_code)]
fn _force_object_use(_: Object) {}
