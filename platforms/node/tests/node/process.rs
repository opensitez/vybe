//! Behaviour tests for `node:process` host imports.
//!
//! Reference: <https://nodejs.org/api/process.html>.
//!
//! Phase 1 covers always-available reads + a few mutators (`cwd`,
//! `chdir`, `exit`-test-deferred). Streams (`stdin`/`stdout`/`stderr`),
//! event listeners (`on('exit', ...)`), and `nextTick` come later.

use vybe_runtime::module_record::ExportEntry;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_proc(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());

    match vm
        .modules
        .get("node:process")
        .and_then(|module| module.exports.get(name))
    {
        Some(ExportEntry::Value(value)) => {
            assert!(args.is_empty(), "node:process.{name} is a value export");
            return value.clone();
        }
        Some(ExportEntry::Function { .. }) => {}
        Some(other) => panic!("unexpected node:process export kind for {name}: {other:?}"),
        None => panic!("missing node:process export: {name}"),
    }

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

    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.modules
        .get("node:process")
        .and_then(|module| module.exports.get(name))
        .is_some()
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
        matches!(
            s.as_str(),
            "darwin" | "linux" | "win32" | "freebsd" | "openbsd"
        ),
        "process.platform() should be Node-recognized, got {}",
        s
    );
}

#[test]
fn arch_returns_recognized_string() {
    let s = as_string(&call_proc("arch", vec![]));
    assert!(
        matches!(
            s.as_str(),
            "x64" | "arm64" | "ia32" | "arm" | "ppc64" | "s390x" | "riscv64"
        ),
        "process.arch() should be Node-recognized, got {}",
        s
    );
}

#[test]
fn version_returns_v_prefixed_string() {
    // Node uses "vMAJOR.MINOR.PATCH" — Vybe surfaces something with a
    // leading 'v' so consuming code's regex still matches.
    let v = as_string(&call_proc("version", vec![]));
    assert!(
        v.starts_with('v'),
        "process.version should start with 'v', got {}",
        v
    );
}

// ── pid / ppid ────────────────────────────────────────────────────

#[test]
fn pid_returns_positive_integer() {
    let v = call_proc("pid", vec![]);
    if let Value::F64(n) = v {
        assert!(
            n > 0.0 && n.is_finite(),
            "process.pid should be positive int, got {}",
            n
        );
        assert_eq!(
            n.fract(),
            0.0,
            "process.pid should be integer-valued, got {}",
            n
        );
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
    let dir = std::env::temp_dir().join(format!("vybe-node-process-chdir-{}", std::process::id()));
    let _ = std::fs::create_dir_all(&dir);

    call_proc("chdir", vec![s(dir.to_str().unwrap())]);
    let after = as_string(&call_proc("cwd", vec![]));
    assert!(
        std::path::Path::new(&after).canonicalize().ok() == dir.canonicalize().ok(),
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
    assert!(
        !names.is_empty(),
        "process.argv should have at least the executable path"
    );
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
        assert!(
            n >= 0.0 && n.is_finite(),
            "uptime should be non-negative finite, got {}",
            n
        );
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
            assert!(
                matches!(&elems[0], Value::F64(_) | Value::I32(_)),
                "hrtime[0] number"
            );
            assert!(
                matches!(&elems[1], Value::F64(_) | Value::I32(_)),
                "hrtime[1] number"
            );
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

// ── versions ─────────────────────────────────────────────────────────────────

#[test]
fn versions_returns_object_with_node_key() {
    let v = call_proc("versions", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("node"),
            "process.versions must have 'node' key"
        );
        return;
    }
    panic!("process.versions expected object, got {:?}", v);
}

#[test]
fn versions_node_is_non_empty_string() {
    let v = call_proc("versions", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::String(s)) = o.properties.get("node") {
            assert!(!s.is_empty(), "process.versions.node must be non-empty");
            return;
        }
        panic!(
            "process.versions.node expected string, got {:?}",
            o.properties.get("node")
        );
    }
    panic!("process.versions expected object, got {:?}", v);
}

// ── ppid ─────────────────────────────────────────────────────────────────────

#[test]
fn ppid_returns_positive_integer() {
    let v = call_proc("ppid", vec![]);
    match v {
        Value::F64(n) => assert!(
            n > 0.0 && n.is_finite(),
            "process.ppid must be positive, got {n}"
        ),
        Value::I32(n) => assert!(n > 0, "process.ppid must be positive, got {n}"),
        Value::I64(n) => assert!(n > 0, "process.ppid must be positive, got {n}"),
        other => panic!("process.ppid expected number, got {:?}", other),
    }
}

// ── title ─────────────────────────────────────────────────────────────────────

#[test]
fn title_returns_string() {
    let v = call_proc("title", vec![]);
    assert!(
        matches!(v, Value::String(_)),
        "process.title must be a string, got {:?}",
        v
    );
}

// ── argv0 ─────────────────────────────────────────────────────────────────────

#[test]
fn argv0_returns_non_empty_string() {
    let v = call_proc("argv0", vec![]);
    let s = as_string(&v);
    assert!(!s.is_empty(), "process.argv0 must be non-empty string");
}

// ── execArgv ──────────────────────────────────────────────────────────────────

#[test]
fn exec_argv_returns_array() {
    let v = call_proc("execArgv", vec![]);
    assert!(
        matches!(v, Value::Object(_)),
        "process.execArgv must be an array"
    );
}

// ── cpuUsage ──────────────────────────────────────────────────────────────────

#[test]
fn cpu_usage_returns_object_with_user_and_system() {
    let v = call_proc("cpuUsage", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("user"),
            "cpuUsage().user must exist"
        );
        assert!(
            o.properties.contains_key("system"),
            "cpuUsage().system must exist"
        );
        return;
    }
    panic!("process.cpuUsage() expected object, got {:?}", v);
}

#[test]
fn cpu_usage_values_are_non_negative() {
    let v = call_proc("cpuUsage", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        for key in ["user", "system"] {
            let val = o.properties.get(key).cloned().unwrap_or(Value::Undefined);
            match val {
                Value::F64(n) => assert!(n >= 0.0, "{key} must be >= 0, got {n}"),
                Value::I32(n) => assert!(n >= 0, "{key} must be >= 0, got {n}"),
                Value::I64(n) => assert!(n >= 0, "{key} must be >= 0, got {n}"),
                _ => panic!("{key} must be a number, got {:?}", val),
            }
        }
        return;
    }
    panic!("process.cpuUsage() expected object");
}

// ── resourceUsage ─────────────────────────────────────────────────────────────

#[test]
fn resource_usage_returns_object() {
    let v = call_proc("resourceUsage", vec![]);
    assert!(
        matches!(v, Value::Object(_)),
        "process.resourceUsage() must return an object"
    );
}

#[test]
fn resource_usage_has_user_cpu_time() {
    let v = call_proc("resourceUsage", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("userCPUTime"),
            "resourceUsage().userCPUTime must exist"
        );
        assert!(
            o.properties.contains_key("systemCPUTime"),
            "resourceUsage().systemCPUTime must exist"
        );
        return;
    }
    panic!("process.resourceUsage() expected object, got {:?}", v);
}

// ── exitCode ──────────────────────────────────────────────────────────────────

#[test]
fn exit_code_is_numeric_or_undefined() {
    let v = call_proc("exitCode", vec![]);
    assert!(
        matches!(
            v,
            Value::I32(_) | Value::I64(_) | Value::F64(_) | Value::Undefined | Value::Null
        ),
        "process.exitCode must be a number or undefined, got {:?}",
        v
    );
}

// ── emitWarning ───────────────────────────────────────────────────────────────

#[test]
fn emit_warning_returns_undefined() {
    let v = call_proc(
        "emitWarning",
        vec![Value::String(std::sync::Arc::from("test warning"))],
    );
    assert!(
        matches!(v, Value::Undefined | Value::Null),
        "process.emitWarning() must return undefined, got {:?}",
        v
    );
}

// ── features ─────────────────────────────────────────────────────────────────

#[test]
fn features_returns_object() {
    let v = call_proc("features", vec![]);
    assert!(
        matches!(v, Value::Object(_)),
        "process.features must be an object"
    );
}

// ── release ───────────────────────────────────────────────────────────────────

#[test]
fn release_returns_object_with_name() {
    let v = call_proc("release", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("name"),
            "process.release.name must exist"
        );
        return;
    }
    panic!("process.release expected object, got {:?}", v);
}

#[test]
fn release_name_is_node() {
    let v = call_proc("release", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::String(name)) = o.properties.get("name") {
            assert_eq!(
                name.as_ref(),
                "node",
                "process.release.name must be 'node', got {name}"
            );
            return;
        }
    }
    panic!("process.release.name expected string 'node'");
}

#[allow(dead_code)]
fn _force_object_use(_: Object) {}

#[test]
fn proposal_node_process_surface_is_registered() {
    let expected = [
        "platform",
        "arch",
        "version",
        "versions",
        "pid",
        "ppid",
        "title",
        "cwd",
        "chdir",
        "exit",
        "argv",
        "argv0",
        "execPath",
        "execArgv",
        "env",
        "uptime",
        "hrtime",
        "memoryUsage",
        "cpuUsage",
        "resourceUsage",
        "exitCode",
        "emitWarning",
        "features",
        "release",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:process imports: {missing:?}"
    );
}
