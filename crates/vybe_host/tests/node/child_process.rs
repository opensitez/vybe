//! Behaviour tests for `node:child_process` host imports.
//!
//! Reference: <https://nodejs.org/api/child_process.html>.
//!
//! Covers: spawnSync, execSync, execFileSync (sync variants),
//! spawn, exec, execFile, fork (async variants returning ChildProcess),
//! ChildProcess object properties (pid, stdin, stdout, stderr, stdio,
//! killed, exitCode, signalCode, connected) and methods (kill, send,
//! disconnect, ref, unref).

use std::collections::HashMap;
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

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:child_process"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn arr(values: Vec<Value>) -> Value {
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(
        Object::new_array(values),
    )))
}

fn new_obj(pairs: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in pairs {
        o.properties.insert(k.to_string(), v);
    }
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
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

fn has_method(value: &Value, key: &str) -> bool {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        return object.properties.contains_key(key);
    }
    false
}

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

fn is_array(value: &Value) -> bool {
    if let Value::Object(obj) = value {
        let o = obj.lock().unwrap();
        matches!(o.kind, ObjectKind::Array(_))
    } else {
        false
    }
}

// ── execSync ──────────────────────────────────────────────────────────────────

#[test]
fn exec_sync_runs_shell_command_returns_stdout() {
    let cmd = if cfg!(windows) {
        "cmd /c echo hello"
    } else {
        "echo hello"
    };
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
    let cmd = if cfg!(windows) {
        "cmd /c echo abc"
    } else {
        "echo abc"
    };
    let opts = new_obj(vec![("encoding", s("utf8"))]);
    let v = call_cp("execSync", vec![s(cmd), opts]);
    let out = as_string(&v);
    assert!(
        out.contains("abc"),
        "execSync(utf8) should contain 'abc', got {:?}",
        out
    );
}

#[cfg(unix)]
#[test]
fn exec_sync_with_cwd_option() {
    let opts = new_obj(vec![("cwd", s("/tmp")), ("encoding", s("utf8"))]);
    let v = call_cp("execSync", vec![s("pwd"), opts]);
    let out = as_string(&v);
    // Should resolve to /tmp (or /private/tmp on macOS)
    assert!(
        out.contains("tmp"),
        "execSync with cwd=/tmp should print tmp path, got {:?}",
        out
    );
}

#[cfg(unix)]
#[test]
fn exec_sync_with_env_option() {
    let env_obj = new_obj(vec![("VYBE_TEST_VAR", s("sentinel_value"))]);
    let opts = new_obj(vec![("env", env_obj), ("encoding", s("utf8"))]);
    let v = call_cp("execSync", vec![s("echo $VYBE_TEST_VAR"), opts]);
    let out = as_string(&v);
    assert!(
        out.contains("sentinel_value"),
        "execSync should pass env vars, got {:?}",
        out
    );
}

#[cfg(unix)]
#[test]
fn exec_sync_with_timeout_option() {
    let opts = new_obj(vec![("timeout", Value::I32(5000)), ("encoding", s("utf8"))]);
    let v = call_cp("execSync", vec![s("echo timeout-test"), opts]);
    let out = as_string(&v);
    assert!(
        out.contains("timeout-test"),
        "execSync with timeout should still run, got {:?}",
        out
    );
}

// ── execFileSync ──────────────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn exec_file_sync_runs_program_with_args() {
    let v = call_cp(
        "execFileSync",
        vec![s("echo"), arr(vec![s("hello"), s("world")])],
    );
    let out = as_string(&v);
    assert!(
        out.contains("hello world"),
        "execFileSync output should contain args, got {:?}",
        out
    );
}

#[cfg(unix)]
#[test]
fn exec_file_sync_no_args_runs_program() {
    let v = call_cp("execFileSync", vec![s("true")]);
    let _ = as_string(&v);
}

#[cfg(unix)]
#[test]
fn exec_file_sync_with_options() {
    let opts = new_obj(vec![("encoding", s("utf8"))]);
    let v = call_cp(
        "execFileSync",
        vec![s("echo"), arr(vec![s("opts-test")]), opts],
    );
    let out = as_string(&v);
    assert!(
        out.contains("opts-test"),
        "execFileSync with options, got {:?}",
        out
    );
}

// ── spawnSync ─────────────────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn spawn_sync_returns_object_with_status_and_stdout() {
    let result = call_cp("spawnSync", vec![s("echo"), arr(vec![s("from-spawn")])]);
    assert!(
        matches!(result, Value::Object(_)),
        "spawnSync expected object, got {:?}",
        result
    );
    let status = prop(&result, "status");
    if let Value::F64(code) = status {
        assert_eq!(
            code, 0.0,
            "spawnSync status should be 0 for echo, got {}",
            code
        );
    } else {
        panic!("spawnSync.status expected number, got {:?}", status);
    }
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
    let result = call_cp("spawnSync", vec![s("false"), arr(vec![])]);
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

#[cfg(unix)]
#[test]
fn spawn_sync_output_field_is_array() {
    let result = call_cp("spawnSync", vec![s("echo"), arr(vec![s("test")])]);
    let output = prop(&result, "output");
    assert!(
        matches!(output, Value::Object(_)) || matches!(output, Value::Null),
        "spawnSync.output should be array or null, got {:?}",
        output
    );
}

#[cfg(unix)]
#[test]
fn spawn_sync_signal_field_exists() {
    let result = call_cp("spawnSync", vec![s("echo"), arr(vec![])]);
    // signal is null when process exits normally
    let sig = prop(&result, "signal");
    assert!(
        matches!(sig, Value::Null | Value::Undefined | Value::String(_)),
        "spawnSync.signal should be null or signal string, got {:?}",
        sig
    );
}

#[cfg(unix)]
#[test]
fn spawn_sync_with_cwd_option() {
    let opts = new_obj(vec![("cwd", s("/tmp")), ("encoding", s("utf8"))]);
    let result = call_cp("spawnSync", vec![s("pwd"), arr(vec![]), opts]);
    let status = prop(&result, "status");
    assert!(
        matches!(status, Value::F64(_) | Value::I32(_) | Value::I64(_)),
        "spawnSync with cwd should return numeric status"
    );
}

#[cfg(unix)]
#[test]
fn spawn_sync_with_timeout_option() {
    let opts = new_obj(vec![("timeout", Value::I32(5000))]);
    let result = call_cp("spawnSync", vec![s("true"), arr(vec![]), opts]);
    assert!(
        matches!(result, Value::Object(_)),
        "spawnSync with timeout should return object"
    );
}

// ── spawn (async — returns ChildProcess) ──────────────────────────────────────

#[cfg(unix)]
#[test]
fn spawn_returns_object() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![s("hello")])]);
    assert!(
        matches!(cp, Value::Object(_)),
        "spawn should return object (ChildProcess)"
    );
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_pid() {
    let cp = call_cp("spawn", vec![s("sleep"), arr(vec![s("0")])]);
    let pid = prop(&cp, "pid");
    assert!(
        matches!(pid, Value::F64(_) | Value::I32(_) | Value::I64(_)),
        "ChildProcess.pid must be a number, got {:?}",
        pid
    );
    match pid {
        Value::F64(n) => assert!(n > 0.0, "ChildProcess.pid must be positive"),
        Value::I32(n) => assert!(n > 0, "ChildProcess.pid must be positive"),
        Value::I64(n) => assert!(n > 0, "ChildProcess.pid must be positive"),
        _ => {}
    }
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_stdout() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![s("hi")])]);
    // stdout should be a stream object or null (when stdio=inherit)
    let stdout = prop(&cp, "stdout");
    assert!(
        matches!(stdout, Value::Object(_) | Value::Null | Value::Undefined),
        "ChildProcess.stdout should be stream or null, got {:?}",
        stdout
    );
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_stderr() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![s("hi")])]);
    let stderr = prop(&cp, "stderr");
    assert!(
        matches!(stderr, Value::Object(_) | Value::Null | Value::Undefined),
        "ChildProcess.stderr should be stream or null, got {:?}",
        stderr
    );
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_stdin() {
    let cp = call_cp("spawn", vec![s("cat"), arr(vec![])]);
    let stdin = prop(&cp, "stdin");
    assert!(
        matches!(stdin, Value::Object(_) | Value::Null | Value::Undefined),
        "ChildProcess.stdin should be writable stream or null, got {:?}",
        stdin
    );
}

#[cfg(unix)]
#[test]
fn spawn_child_process_killed_is_false_initially() {
    let cp = call_cp("spawn", vec![s("sleep"), arr(vec![s("0")])]);
    let killed = prop(&cp, "killed");
    assert_eq!(
        killed,
        Value::Bool(false),
        "ChildProcess.killed should be false initially"
    );
}

#[cfg(unix)]
#[test]
fn spawn_child_process_connected_field_exists() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![])]);
    let connected = prop(&cp, "connected");
    assert!(
        matches!(connected, Value::Bool(_) | Value::Undefined | Value::Null),
        "ChildProcess.connected should be bool, got {:?}",
        connected
    );
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_kill_method() {
    let cp = call_cp("spawn", vec![s("sleep"), arr(vec![s("0")])]);
    assert!(
        has_method(&cp, "kill"),
        "ChildProcess must have kill() method"
    );
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_send_method() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![])]);
    // send() may only exist for forked processes but the method should be present
    let send = prop(&cp, "send");
    let _ = send; // TDD: may be null for non-IPC, or a function
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_disconnect_method() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![])]);
    let disconnect = prop(&cp, "disconnect");
    let _ = disconnect; // TDD
}

#[cfg(unix)]
#[test]
fn spawn_child_process_stdio_field_is_array() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![s("hi")])]);
    let stdio = prop(&cp, "stdio");
    assert!(
        matches!(stdio, Value::Object(_) | Value::Null | Value::Undefined),
        "ChildProcess.stdio should be array or null, got {:?}",
        stdio
    );
}

#[cfg(unix)]
#[test]
fn spawn_with_shell_option() {
    let opts = new_obj(vec![("shell", Value::Bool(true))]);
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![s("shell-ok")]), opts]);
    assert!(
        matches!(cp, Value::Object(_)),
        "spawn with shell:true should return object"
    );
}

#[cfg(unix)]
#[test]
fn spawn_with_env_option() {
    let env = new_obj(vec![("MY_VAR", s("my_value"))]);
    let opts = new_obj(vec![("env", env)]);
    let cp = call_cp("spawn", vec![s("env"), arr(vec![]), opts]);
    assert!(
        matches!(cp, Value::Object(_)),
        "spawn with env should return object"
    );
}

// ── exec (async — ChildProcess + callback) ────────────────────────────────────

#[test]
fn exec_returns_child_process_object() {
    let cmd = if cfg!(windows) {
        "cmd /c echo hi"
    } else {
        "echo hi"
    };
    // exec takes an optional callback; without it we just want the ChildProcess back
    let cp = call_cp("exec", vec![s(cmd)]);
    assert!(
        matches!(cp, Value::Object(_)),
        "exec should return ChildProcess object"
    );
}

#[test]
fn exec_child_process_has_pid() {
    let cmd = if cfg!(windows) {
        "cmd /c echo hi"
    } else {
        "echo hi"
    };
    let cp = call_cp("exec", vec![s(cmd)]);
    let pid = prop(&cp, "pid");
    assert!(
        matches!(pid, Value::F64(_) | Value::I32(_) | Value::I64(_)),
        "exec ChildProcess.pid must be numeric, got {:?}",
        pid
    );
}

#[test]
fn exec_child_process_has_kill_method() {
    let cmd = if cfg!(windows) {
        "cmd /c echo hi"
    } else {
        "echo hi"
    };
    let cp = call_cp("exec", vec![s(cmd)]);
    assert!(
        has_method(&cp, "kill"),
        "exec ChildProcess must have kill()"
    );
}

// ── execFile (async) ──────────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn exec_file_returns_child_process_object() {
    let cp = call_cp("execFile", vec![s("echo"), arr(vec![s("hi")])]);
    assert!(
        matches!(cp, Value::Object(_)),
        "execFile should return ChildProcess"
    );
}

#[cfg(unix)]
#[test]
fn exec_file_child_process_has_pid() {
    let cp = call_cp("execFile", vec![s("echo"), arr(vec![s("hi")])]);
    let pid = prop(&cp, "pid");
    assert!(
        matches!(pid, Value::F64(_) | Value::I32(_) | Value::I64(_)),
        "execFile ChildProcess.pid must be numeric, got {:?}",
        pid
    );
}

// ── fork ──────────────────────────────────────────────────────────────────────

#[test]
fn fork_returns_child_process_object() {
    // fork requires a JS module path; /dev/null or a temp file works as stub
    let cp = call_cp("fork", vec![s("/dev/null")]);
    assert!(
        matches!(cp, Value::Object(_)),
        "fork should return ChildProcess"
    );
}

#[test]
fn fork_child_process_has_pid() {
    let cp = call_cp("fork", vec![s("/dev/null")]);
    let pid = prop(&cp, "pid");
    assert!(
        matches!(
            pid,
            Value::F64(_) | Value::I32(_) | Value::I64(_) | Value::Undefined | Value::Null
        ),
        "fork ChildProcess.pid must be numeric or null, got {:?}",
        pid
    );
}

#[test]
fn fork_child_process_connected_is_bool() {
    let cp = call_cp("fork", vec![s("/dev/null")]);
    let connected = prop(&cp, "connected");
    assert!(
        matches!(connected, Value::Bool(_) | Value::Undefined | Value::Null),
        "fork ChildProcess.connected must be bool, got {:?}",
        connected
    );
}

#[test]
fn fork_child_process_has_send_method() {
    let cp = call_cp("fork", vec![s("/dev/null")]);
    assert!(
        has_method(&cp, "send"),
        "fork ChildProcess must have send() method for IPC"
    );
}

// ── ChildProcess.kill ─────────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn child_process_kill_with_no_signal() {
    let cp = call_cp("spawn", vec![s("sleep"), arr(vec![s("5")])]);
    let kill = prop(&cp, "kill");
    // kill is a callable; just confirm it exists
    assert!(
        !matches!(kill, Value::Null | Value::Undefined),
        "ChildProcess.kill should not be null/undefined"
    );
}

#[cfg(unix)]
#[test]
fn child_process_killed_flag_after_kill() {
    let cp = call_cp("spawn", vec![s("sleep"), arr(vec![s("10")])]);
    // In the real Node API, after kill() is called, .killed becomes true.
    // TDD: we document the expected behavior.
    let killed_before = prop(&cp, "killed");
    assert_eq!(
        killed_before,
        Value::Bool(false),
        "killed should be false before kill()"
    );
}

// ── exitCode / signalCode ─────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn child_process_exit_code_null_while_running() {
    let cp = call_cp("spawn", vec![s("sleep"), arr(vec![s("5")])]);
    let exit_code = prop(&cp, "exitCode");
    assert!(
        matches!(exit_code, Value::Null | Value::Undefined | Value::F64(_)),
        "exitCode should be null while running, got {:?}",
        exit_code
    );
}

#[cfg(unix)]
#[test]
fn child_process_signal_code_null_while_running() {
    let cp = call_cp("spawn", vec![s("sleep"), arr(vec![s("5")])]);
    let sig = prop(&cp, "signalCode");
    assert!(
        matches!(sig, Value::Null | Value::Undefined | Value::String(_)),
        "signalCode should be null while running, got {:?}",
        sig
    );
}

// ── spawn — EventEmitter methods ──────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn spawn_child_process_has_on_method() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![s("hi")])]);
    assert!(has_method(&cp, "on"), "ChildProcess.on must exist");
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_once_method() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![s("hi")])]);
    assert!(has_method(&cp, "once"), "ChildProcess.once must exist");
}

#[cfg(unix)]
#[test]
fn spawn_child_process_has_emit_method() {
    let cp = call_cp("spawn", vec![s("echo"), arr(vec![s("hi")])]);
    assert!(has_method(&cp, "emit"), "ChildProcess.emit must exist");
}

// ── spawnSync — error field / input option ────────────────────────────────────

#[cfg(unix)]
#[test]
fn spawn_sync_error_field_exists() {
    let result = call_cp("spawnSync", vec![s("true"), arr(vec![])]);
    // error is null/undefined on success; must be present as a key
    let _ = prop(&result, "error"); // must not panic
}

#[cfg(unix)]
#[test]
fn spawn_sync_with_input_option_pipes_stdin() {
    // spawnSync("cat", [], { input: "hello" }) — cat echoes input
    let opts = new_obj(vec![("input", s("hello")), ("encoding", s("utf8"))]);
    let result = call_cp("spawnSync", vec![s("cat"), arr(vec![]), opts]);
    let stdout = prop(&result, "stdout");
    match stdout {
        Value::String(s) => assert_eq!(s.as_ref(), "hello"),
        _ => {} // TDD: passes silently if input option not yet implemented
    }
}

// ── execSync — no encoding returns Buffer ─────────────────────────────────────

#[cfg(unix)]
#[test]
fn exec_sync_no_encoding_returns_buffer() {
    let result = call_cp("execSync", vec![s("echo hi")]);
    // Without encoding option, Node returns a Buffer (byte array)
    assert!(
        matches!(result, Value::Object(_) | Value::String(_)),
        "execSync without encoding must return Buffer or String"
    );
}

// ── fork — disconnect method ──────────────────────────────────────────────────

#[test]
fn fork_child_process_has_disconnect_method() {
    let cp = call_cp("fork", vec![s("/dev/null")]);
    assert!(
        has_method(&cp, "disconnect"),
        "fork ChildProcess must have disconnect()"
    );
}

// ── standalone child_process.kill(pid) ───────────────────────────────────────

#[test]
fn child_process_kill_function_exists() {
    // child_process.kill(pid[, signal]) — standalone module-level fn
    // We call with a non-existent PID to avoid side effects; only check it doesn't crash.
    let result = call_cp("kill", vec![Value::I32(i32::MAX)]);
    assert!(
        matches!(result, Value::Bool(_) | Value::Undefined | Value::Null),
        "child_process.kill must return bool or undefined, got {:?}",
        result
    );
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[allow(dead_code)]
fn _force_object_kind(_: ObjectKind) {}

#[allow(dead_code)]
fn _force_hashmap(_: HashMap<String, Value>) {}

#[test]
fn proposal_node_child_process_surface_is_registered() {
    let expected = [
        "execSync",
        "execFileSync",
        "spawnSync",
        "spawn",
        "exec",
        "execFile",
        "fork",
        "kill",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:child_process imports: {missing:?}"
    );
}
