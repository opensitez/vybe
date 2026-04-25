//! Behaviour tests for `node:child_process` host imports.
//!
//! Reference: <https://nodejs.org/api/child_process.html>.
//!
//! Phase 1 covers sync variants only: `spawnSync`, `execSync`,
//! `execFileSync`. Async (callback) and stream (`spawn`/`exec` w/
//! pipes) variants come later.
//!
//! Tests use `echo`/`cat`/`ls` on Unix and `cmd /c echo` on Windows.

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_cp(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-child-process-test>");
    let import_idx = chunk.add_import("node:child_process", name);
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

fn arr(values: Vec<Value>) -> Value {
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object::new_array(values))))
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

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

// ── execSync ──────────────────────────────────────────────────────

#[test]
fn exec_sync_runs_shell_command_returns_stdout() {
    // Node `execSync(command)` runs the command via `/bin/sh -c` (Unix)
    // or `cmd.exe /d /s /c` (Windows) and returns stdout as a Buffer
    // (or string with `encoding` option).
    let cmd = if cfg!(windows) { "cmd /c echo hello" } else { "echo hello" };
    let v = call_cp("execSync", vec![s(cmd)]);
    let out = as_string(&v);
    assert!(
        out.trim_end() == "hello" || out.contains("hello"),
        "execSync output should contain 'hello', got {:?}",
        out
    );
}

#[test]
fn exec_sync_with_options_uses_encoding() {
    // Node: `execSync(cmd, { encoding: 'utf8' })` returns a string.
    let cmd = if cfg!(windows) { "cmd /c echo abc" } else { "echo abc" };
    let opts = {
        let mut o = Object::new();
        o.properties.insert("encoding".into(), s("utf8"));
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
    };
    let v = call_cp("execSync", vec![s(cmd), opts]);
    let out = as_string(&v);
    assert!(out.contains("abc"), "execSync(utf8) should contain 'abc', got {:?}", out);
}

// ── execFileSync ──────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn exec_file_sync_runs_program_with_args() {
    let v = call_cp(
        "execFileSync",
        vec![s("echo"), arr(vec![s("hello"), s("world")])],
    );
    let out = as_string(&v);
    assert!(out.contains("hello world"), "execFileSync output should contain args, got {:?}", out);
}

#[cfg(unix)]
#[test]
fn exec_file_sync_no_args_runs_program() {
    let v = call_cp("execFileSync", vec![s("true")]);
    // `true` produces no output but the call should succeed.
    let _ = as_string(&v);
}

// ── spawnSync ─────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn spawn_sync_returns_object_with_status_and_stdout() {
    // Node spawnSync returns:
    //   { pid, output, stdout, stderr, status, signal, error }
    let result = call_cp(
        "spawnSync",
        vec![s("echo"), arr(vec![s("from-spawn")])],
    );
    // Expect an object.
    assert!(matches!(result, Value::Object(_)), "spawnSync expected object, got {:?}", result);

    // status should be 0 (success).
    let status = prop(&result, "status");
    if let Value::F64(code) = status {
        assert_eq!(code, 0.0, "spawnSync status should be 0 for echo, got {}", code);
    } else {
        panic!("spawnSync.status expected number, got {:?}", status);
    }

    // stdout should contain our argument.
    let stdout = as_string(&prop(&result, "stdout"));
    assert!(
        stdout.contains("from-spawn"),
        "spawnSync.stdout should contain 'from-spawn', got {:?}",
        stdout
    );
}

#[cfg(unix)]
#[test]
fn spawn_sync_failed_command_status_non_zero() {
    let result = call_cp(
        "spawnSync",
        vec![s("false"), arr(vec![])],
    );
    let status = prop(&result, "status");
    if let Value::F64(code) = status {
        assert_ne!(code, 0.0, "false should have non-zero status, got {}", code);
    } else {
        panic!("spawnSync.status expected number, got {:?}", status);
    }
}

#[cfg(unix)]
#[test]
fn spawn_sync_returns_pid() {
    let result = call_cp("spawnSync", vec![s("true"), arr(vec![])]);
    let pid = prop(&result, "pid");
    if let Value::F64(n) = pid {
        assert!(n > 0.0, "spawnSync.pid should be positive, got {}", n);
    } else {
        panic!("spawnSync.pid expected number, got {:?}", pid);
    }
}

#[cfg(unix)]
#[test]
fn spawn_sync_captures_stderr() {
    // sh -c writes to stderr explicitly so we have a deterministic
    // way to verify stderr is captured.
    let result = call_cp(
        "spawnSync",
        vec![s("sh"), arr(vec![s("-c"), s("echo errmsg 1>&2")])],
    );
    let stderr = as_string(&prop(&result, "stderr"));
    assert!(
        stderr.contains("errmsg"),
        "spawnSync.stderr should contain 'errmsg', got {:?}",
        stderr
    );
}

#[allow(dead_code)]
fn _force_object_kind(_: ObjectKind) {}
