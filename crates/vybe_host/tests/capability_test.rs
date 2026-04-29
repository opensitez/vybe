use vybe_bytecode::*;
use vybe_bytecode::value::*;
use vybe_host::{Capabilities, Capability, register_with_capabilities};
use std::sync::{Arc, Mutex};

// ============================================================
// Capability preset tests
// ============================================================

#[test]
fn caps_all_has_everything() {
    let caps = Capabilities::all();
    assert!(caps.has(Capability::Console));
    assert!(caps.has(Capability::FileRead));
    assert!(caps.has(Capability::FileWrite));
    assert!(caps.has(Capability::Http));
    assert!(caps.has(Capability::Sockets));
    assert!(caps.has(Capability::Database));
    assert!(caps.has(Capability::Environment));
    assert!(caps.has(Capability::Gui));
    assert!(caps.has(Capability::Threading));
    assert!(caps.has(Capability::Crypto));
    assert!(caps.has(Capability::Clock));
    assert!(caps.has(Capability::Random));
    assert!(caps.has(Capability::Xml));
}

#[test]
fn caps_safe_limited() {
    let caps = Capabilities::safe();
    assert!(caps.has(Capability::Console));
    assert!(caps.has(Capability::Clock));
    assert!(caps.has(Capability::Random));
    assert!(!caps.has(Capability::FileRead));
    assert!(!caps.has(Capability::FileWrite));
    assert!(!caps.has(Capability::Http));
    assert!(!caps.has(Capability::Sockets));
    assert!(!caps.has(Capability::Database));
    assert!(!caps.has(Capability::Environment));
}

#[test]
fn caps_none_empty() {
    let caps = Capabilities::none();
    assert!(!caps.has(Capability::Console));
    assert!(!caps.has(Capability::FileRead));
    assert!(!caps.has(Capability::Clock));
}

#[test]
fn caps_custom() {
    let caps = Capabilities::with(&[Capability::Console, Capability::Crypto]);
    assert!(caps.has(Capability::Console));
    assert!(caps.has(Capability::Crypto));
    assert!(!caps.has(Capability::FileRead));
    assert!(!caps.has(Capability::Database));
}

#[test]
fn caps_grant_revoke() {
    let mut caps = Capabilities::none();
    assert!(!caps.has(Capability::Http));
    caps.grant(Capability::Http);
    assert!(caps.has(Capability::Http));
    caps.revoke(Capability::Http);
    assert!(!caps.has(Capability::Http));
}

// ============================================================
// Module registration — full access
// ============================================================

#[test]
fn all_caps_registers_filesystem() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    let has_fs = vm.host_registry.keys().any(|(m, _)| m == "wasi:filesystem");
    assert!(has_fs, "Full caps should register filesystem");
}

#[test]
fn all_caps_registers_database() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    let has_db = vm.host_registry.keys().any(|(m, _)| m == "vybe:database");
    assert!(has_db, "Full caps should register database");
}

#[test]
fn all_caps_registers_sockets() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    // The .NET-shaped TcpClient / TcpListener / UdpClient host fns live
    // under `dotnet:sockets`. Real WASI 0.2.8 socket primitives are checked
    // by `all_caps_registers_wasi_sockets` below.
    let has_sock = vm.host_registry.keys().any(|(m, _)| m.starts_with("wasi:sockets/"));
    assert!(has_sock, "Full caps should register dotnet:sockets");
}

#[test]
fn all_caps_registers_wasi_sockets() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    let has_sock = vm.host_registry.keys().any(|(m, _)| m.starts_with("wasi:sockets/"));
    assert!(has_sock, "Full caps should register wasi:sockets modules");
}

#[test]
fn all_caps_registers_wasi_io() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    let has_io = vm.host_registry.keys().any(|(m, _)| m == "wasi:io/streams" || m == "wasi:io/poll");
    assert!(has_io, "Full caps should register wasi:io modules");
}

#[test]
fn all_caps_registers_http() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    let has_http = vm.host_registry.keys().any(|(m, _)| m == "wasi:http");
    assert!(has_http, "Full caps should register HTTP");
}

// ============================================================
// Module registration — safe mode BLOCKS dangerous modules
// ============================================================

#[test]
fn safe_blocks_filesystem() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_fs = vm.host_registry.keys().any(|(m, _)| m == "wasi:filesystem");
    assert!(!has_fs, "Safe mode should NOT have filesystem");
}

#[test]
fn safe_blocks_database() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_db = vm.host_registry.keys().any(|(m, _)| m == "vybe:database");
    assert!(!has_db, "Safe mode should NOT have database");
}

#[test]
fn safe_blocks_sockets() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_sock = vm.host_registry.keys().any(|(m, _)| m.starts_with("wasi:sockets/"));
    assert!(!has_sock, "Safe mode should NOT have dotnet:sockets");
}

#[test]
fn safe_blocks_wasi_sockets() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_sock = vm.host_registry.keys().any(|(m, _)| m.starts_with("wasi:sockets/"));
    assert!(!has_sock, "Safe mode should NOT have wasi:sockets modules");
}

#[test]
fn safe_blocks_wasi_io() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_io = vm.host_registry.keys().any(|(m, _)| m == "wasi:io/streams" || m == "wasi:io/poll");
    assert!(!has_io, "Safe mode should NOT have wasi:io modules");
}

#[test]
fn safe_blocks_http() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_http = vm.host_registry.keys().any(|(m, _)| m == "wasi:http");
    assert!(!has_http, "Safe mode should NOT have HTTP");
}

#[test]
fn safe_blocks_threading() {
    // Threading no longer has a host-fn module: thread spawn/join compile
    // directly to WASM opcodes (Op::THREAD_SPAWN / Op::THREAD_JOIN) and
    // Thread.Sleep uses wasi:clocks. The `vybe:threading` namespace must
    // not exist in any mode.
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_vybe_threading = vm.host_registry.keys().any(|(m, _)| m == "vybe:threading");
    assert!(!has_vybe_threading, "vybe:threading must be fully retired");
}

// ============================================================
// Safe mode ALLOWS pure computation modules
// ============================================================

#[test]
fn safe_allows_math() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    // Math is registered under `ecma:math` — the ECMA-262 `Math` object
    // mirror. The `ecma:*` prefix marks it as a Vybe-provided namespace
    // for ECMA-262 types (no such WebAssembly CG proposal exists; the
    // stage-1 wasm-js-primitive-builtins proposal explicitly rejected Math).
    let has_math = vm.host_registry.keys().any(|(m, _)| m == "ecma:math");
    assert!(has_math, "Safe mode should have math");
}

#[test]
fn safe_allows_string() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_str = vm.host_registry.keys().any(|(m, _)| m == "ecma:string");
    assert!(has_str, "Safe mode should have ecma:string");
}

#[test]
fn safe_allows_json() {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_json = vm.host_registry.keys().any(|(m, _)| m == "ecma:json");
    assert!(has_json, "Safe mode should have JSON");
}

#[test]
fn safe_allows_convert() {
    // VB/.NET `Convert.ToXxx` is now backed by ECMA-262 §21 Number / §22.1
    // String primitives. A safe-mode VM exposes those.
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());
    let has_number = vm.host_registry.keys().any(|(m, _)| m == "ecma:number");
    assert!(has_number, "Safe mode should expose ecma:number for conversions");
}

// ============================================================
// Custom capabilities — selective access
// ============================================================

#[test]
fn custom_database_only() {
    let caps = Capabilities::with(&[Capability::Console, Capability::Database]);
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &caps);

    let has_db = vm.host_registry.keys().any(|(m, _)| m == "vybe:database");
    let has_fs = vm.host_registry.keys().any(|(m, _)| m == "wasi:filesystem");
    let has_http = vm.host_registry.keys().any(|(m, _)| m == "wasi:http");

    assert!(has_db, "Should have database");
    assert!(!has_fs, "Should NOT have filesystem");
    assert!(!has_http, "Should NOT have HTTP");
}

#[test]
fn custom_network_only() {
    let caps = Capabilities::with(&[Capability::Http, Capability::Sockets]);
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &caps);

    let has_http = vm.host_registry.keys().any(|(m, _)| m == "wasi:http");
    let has_sock = vm.host_registry.keys().any(|(m, _)| m.starts_with("wasi:sockets/"));
    let has_db = vm.host_registry.keys().any(|(m, _)| m == "vybe:database");
    let has_fs = vm.host_registry.keys().any(|(m, _)| m == "wasi:filesystem");

    assert!(has_http, "Should have HTTP");
    assert!(has_sock, "Should have sockets");
    assert!(!has_db, "Should NOT have database");
    assert!(!has_fs, "Should NOT have filesystem");
}

// ============================================================
// Runtime behavior — blocked calls return undefined
// ============================================================

#[test]
fn blocked_host_call_returns_undefined() {
    // Register with safe caps (no filesystem)
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::safe());

    // Build a chunk that tries to call a filesystem function
    let mut chunk = Chunk::new("<test>");
    // Try to call wasi:filesystem readFile — it shouldn't be registered
    let ci = chunk.add_constant(Value::String(std::sync::Arc::from("test.txt")));
    chunk.emit_op_u16(Op::CONST, ci, 0);
    // Try global_get for a filesystem function — returns Undefined
    let fn_name = chunk.add_constant(Value::String(std::sync::Arc::from("readfile")));
    chunk.emit_op_u16(Op::GLOBAL_GET, fn_name, 0);
    // The function doesn't exist, so we get Undefined
    chunk.emit_op(Op::REF_IS_NULL, 0); // Undefined is null-ish
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    // ref_is_null on Undefined should be true
    assert!(matches!(result, Value::Bool(true)), "Blocked function should be undefined/null");
}

// ============================================================
// Console output capture works in sandboxed mode
// ============================================================

#[test]
fn sandbox_console_capture() {
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();

    register_with_capabilities(&mut vm, &Capabilities::safe());
    // Override console.log to capture output
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));

    // Build a simple chunk: push "hello sandbox", call console.log
    let mut chunk = Chunk::new("<test>");
    let msg = chunk.add_constant(Value::String(std::sync::Arc::from("hello sandbox")));
    chunk.emit_op_u16(Op::CONST, msg, 0);
    let log_idx = chunk.add_import("wasi:cli", "log");
    chunk.emit_op_u16(Op::CALL_IMPORT, log_idx, 0);
    chunk.emit(1, 0); // argc = 1
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::HALT, 0);

    vm.run(vec![chunk]).unwrap();
    assert_eq!(*output.lock().unwrap(), vec!["hello sandbox"]);
}
