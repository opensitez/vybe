//! Behaviour tests for `node:os` host imports.
//!
//! Reference: <https://nodejs.org/api/os.html>.
//!
//! These tests assert what we can verify portably (return shapes,
//! non-empty results) rather than exact values that vary per host.

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_os(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-os-test>");
    let import_idx = chunk.add_import("node:os", name);
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

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

// ── platform / arch / type / release / version ────────────────────

#[test]
fn platform_returns_recognized_string() {
    let v = call_os("platform", vec![]);
    let s = as_string(&v);
    assert!(
        matches!(s.as_str(), "darwin" | "linux" | "win32" | "freebsd" | "openbsd" | "sunos" | "aix"),
        "platform() should return a Node-recognized string, got {}",
        s
    );
}

#[test]
fn arch_returns_recognized_string() {
    let v = call_os("arch", vec![]);
    let s = as_string(&v);
    assert!(
        matches!(
            s.as_str(),
            "x64" | "arm64" | "ia32" | "arm" | "ppc64" | "s390x" | "mips" | "mipsel" | "riscv64"
        ),
        "arch() should return a Node-recognized string, got {}",
        s
    );
}

#[test]
fn type_returns_non_empty_string() {
    let v = call_os("type", vec![]);
    let s = as_string(&v);
    assert!(!s.is_empty(), "type() should return non-empty string");
}

// ── hostname / tmpdir / homedir ───────────────────────────────────

#[test]
fn hostname_returns_non_empty_string() {
    let v = call_os("hostname", vec![]);
    let s = as_string(&v);
    assert!(!s.is_empty(), "hostname() should return non-empty string");
}

#[test]
fn tmpdir_returns_existing_directory() {
    let v = call_os("tmpdir", vec![]);
    let s = as_string(&v);
    assert!(
        std::path::Path::new(&s).is_dir(),
        "tmpdir() should point to an existing directory, got {}",
        s
    );
}

#[test]
fn homedir_returns_existing_directory() {
    let v = call_os("homedir", vec![]);
    let s = as_string(&v);
    assert!(
        std::path::Path::new(&s).is_dir(),
        "homedir() should point to an existing directory, got {}",
        s
    );
}

// ── endianness / EOL ──────────────────────────────────────────────

#[test]
fn endianness_returns_be_or_le() {
    let v = call_os("endianness", vec![]);
    let s = as_string(&v);
    assert!(s == "BE" || s == "LE", "endianness() should be BE or LE, got {}", s);
}

#[test]
fn eol_returns_unix_or_windows_newline() {
    // EOL is a property in Node, exposed here as a 0-arg fn.
    let v = call_os("EOL", vec![]);
    let s = as_string(&v);
    assert!(s == "\n" || s == "\r\n", "EOL should be \\n or \\r\\n, got {:?}", s);
}

// ── totalmem / freemem / uptime ───────────────────────────────────

#[test]
fn totalmem_returns_positive_number() {
    let v = call_os("totalmem", vec![]);
    if let Value::F64(n) = v {
        assert!(n > 0.0 && n.is_finite(), "totalmem() should be positive finite, got {}", n);
    } else {
        panic!("totalmem() expected number, got {:?}", v);
    }
}

#[test]
fn freemem_returns_non_negative_number() {
    let v = call_os("freemem", vec![]);
    if let Value::F64(n) = v {
        assert!(n >= 0.0 && n.is_finite(), "freemem() should be non-negative finite, got {}", n);
    } else {
        panic!("freemem() expected number, got {:?}", v);
    }
}

#[test]
fn uptime_returns_non_negative_number() {
    let v = call_os("uptime", vec![]);
    if let Value::F64(n) = v {
        assert!(n >= 0.0 && n.is_finite(), "uptime() should be non-negative finite, got {}", n);
    } else {
        panic!("uptime() expected number, got {:?}", v);
    }
}

// ── cpus / userInfo ───────────────────────────────────────────────

#[test]
fn cpus_returns_non_empty_array() {
    let v = call_os("cpus", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            assert!(!elems.is_empty(), "cpus() should return at least one CPU");
            // Each CPU should be an object with at least `model` and `speed`.
            if let Value::Object(first) = &elems[0] {
                let f = first.lock().unwrap();
                assert!(f.properties.contains_key("model"), "cpu[0].model expected");
                assert!(f.properties.contains_key("speed"), "cpu[0].speed expected");
            } else {
                panic!("cpu[0] expected object, got {:?}", elems[0]);
            }
            return;
        }
    }
    panic!("cpus() expected array, got {:?}", v);
}

#[test]
fn user_info_returns_object_with_username() {
    let v = call_os("userInfo", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        let username = o.properties.get("username").cloned().unwrap_or(Value::Null);
        if let Value::String(text) = username {
            assert!(!text.is_empty(), "userInfo().username should be non-empty");
            return;
        }
        panic!("userInfo().username expected string, got {:?}", username);
    }
    panic!("userInfo() expected object, got {:?}", v);
}

// Suppress unused-import warning in builds where Object isn't used yet.
#[allow(dead_code)]
fn _force_object_use(_: Object) {}
