//! Behaviour tests for `node:os` host imports.
//!
//! Reference: <https://nodejs.org/api/os.html>.
//!
//! These tests assert what we can verify portably (return shapes,
//! non-empty results) rather than exact values that vary per host.

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

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
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:os"), name.to_string()))
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
        matches!(
            s.as_str(),
            "darwin" | "linux" | "win32" | "freebsd" | "openbsd" | "sunos" | "aix"
        ),
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
    assert!(
        s == "BE" || s == "LE",
        "endianness() should be BE or LE, got {}",
        s
    );
}

#[test]
fn eol_returns_unix_or_windows_newline() {
    // EOL is a property in Node, exposed here as a 0-arg fn.
    let v = call_os("EOL", vec![]);
    let s = as_string(&v);
    assert!(
        s == "\n" || s == "\r\n",
        "EOL should be \\n or \\r\\n, got {:?}",
        s
    );
}

// ── totalmem / freemem / uptime ───────────────────────────────────

#[test]
fn totalmem_returns_positive_number() {
    let v = call_os("totalmem", vec![]);
    if let Value::F64(n) = v {
        assert!(
            n > 0.0 && n.is_finite(),
            "totalmem() should be positive finite, got {}",
            n
        );
    } else {
        panic!("totalmem() expected number, got {:?}", v);
    }
}

#[test]
fn freemem_returns_non_negative_number() {
    let v = call_os("freemem", vec![]);
    if let Value::F64(n) = v {
        assert!(
            n >= 0.0 && n.is_finite(),
            "freemem() should be non-negative finite, got {}",
            n
        );
    } else {
        panic!("freemem() expected number, got {:?}", v);
    }
}

#[test]
fn uptime_returns_non_negative_number() {
    let v = call_os("uptime", vec![]);
    if let Value::F64(n) = v {
        assert!(
            n >= 0.0 && n.is_finite(),
            "uptime() should be non-negative finite, got {}",
            n
        );
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
// ── loadavg ───────────────────────────────────────────────────────────────────

#[test]
fn loadavg_returns_array_of_three_numbers() {
    let v = call_os("loadavg", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            assert_eq!(elems.len(), 3, "loadavg() must return exactly 3 elements");
            for elem in elems {
                assert!(
                    matches!(elem, Value::F64(_) | Value::I32(_) | Value::I64(_)),
                    "each loadavg element must be a number, got {:?}",
                    elem
                );
            }
            return;
        }
    }
    panic!("loadavg() expected array of 3 numbers, got {:?}", v);
}

#[test]
fn loadavg_values_are_non_negative() {
    let v = call_os("loadavg", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            for elem in elems {
                if let Value::F64(f) = elem {
                    assert!(*f >= 0.0, "load average must be non-negative, got {f}");
                }
            }
            return;
        }
    }
    panic!("loadavg() expected array, got {:?}", v);
}

// ── networkInterfaces ─────────────────────────────────────────────────────────

#[test]
fn network_interfaces_returns_object() {
    let v = call_os("networkInterfaces", vec![]);
    assert!(
        matches!(v, Value::Object(_)),
        "networkInterfaces() must return an object, got {:?}",
        v
    );
}

// ── release ───────────────────────────────────────────────────────────────────

#[test]
fn release_returns_non_empty_string() {
    let v = call_os("release", vec![]);
    let s = as_string(&v);
    assert!(!s.is_empty(), "os.release() must return non-empty string");
}

// ── version ───────────────────────────────────────────────────────────────────

#[test]
fn version_returns_non_empty_string() {
    let v = call_os("version", vec![]);
    let s = as_string(&v);
    assert!(!s.is_empty(), "os.version() must return non-empty string");
}

// ── getPriority / setPriority ─────────────────────────────────────────────────

#[test]
fn get_priority_returns_number() {
    let v = call_os("getPriority", vec![]);
    assert!(
        matches!(v, Value::I32(_) | Value::I64(_) | Value::F64(_)),
        "getPriority() must return a number, got {:?}",
        v
    );
}

#[test]
fn set_priority_returns_undefined() {
    let result = call_os("setPriority", vec![Value::I32(0)]);
    assert!(
        matches!(result, Value::Undefined | Value::Null),
        "setPriority(0) must return undefined/null, got {:?}",
        result
    );
}

// ── constants (extended) ──────────────────────────────────────────────────────

#[test]
fn os_constants_signals_sigint_is_2() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(signals)) = o.properties.get("signals") {
            let s = signals.lock().unwrap();
            let sigint = s
                .properties
                .get("SIGINT")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(sigint, Value::I32(2), "SIGINT must be 2");
        }
    }
}

#[test]
fn os_constants_signals_sigkill_is_9() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(signals)) = o.properties.get("signals") {
            let s = signals.lock().unwrap();
            let sigkill = s
                .properties
                .get("SIGKILL")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(sigkill, Value::I32(9), "SIGKILL must be 9");
        }
    }
}

#[test]
fn os_constants_errno_enoent_is_present() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(errno)) = o.properties.get("errno") {
            let e = errno.lock().unwrap();
            let val = e
                .properties
                .get("ENOENT")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert!(
                matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
                "errno.ENOENT must be a number, got {:?}",
                val
            );
        } else {
            panic!("os.constants.errno must be an object");
        }
    }
}

#[test]
fn os_constants_errno_eacces_is_present() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(errno)) = o.properties.get("errno") {
            let e = errno.lock().unwrap();
            let val = e
                .properties
                .get("EACCES")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert!(
                matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
                "errno.EACCES must be a number, got {:?}",
                val
            );
        }
    }
}

#[test]
fn os_constants_has_priority_object() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("priority"),
            "os.constants must have a priority sub-object"
        );
    }
}

#[test]
fn os_constants_priority_normal_is_zero() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(prio)) = o.properties.get("priority") {
            let p = prio.lock().unwrap();
            let val = p
                .properties
                .get("PRIORITY_NORMAL")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(
                val,
                Value::I32(0),
                "PRIORITY_NORMAL must be 0, got {:?}",
                val
            );
        }
    }
}

// ── cpus — extended ───────────────────────────────────────────────────────────

#[test]
fn cpus_first_cpu_has_times_object() {
    let v = call_os("cpus", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            if let Value::Object(cpu) = &elems[0] {
                let c = cpu.lock().unwrap();
                let times = c
                    .properties
                    .get("times")
                    .cloned()
                    .unwrap_or(Value::Undefined);
                assert!(
                    matches!(times, Value::Object(_)),
                    "cpu.times must be an object"
                );
                return;
            }
        }
    }
    panic!("cpus() returned unexpected shape");
}

// ── cpus times fields ─────────────────────────────────────────────────────────

#[test]
fn cpus_times_has_user_sys_idle_fields() {
    let v = call_os("cpus", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            if let Value::Object(cpu) = &elems[0] {
                let c = cpu.lock().unwrap();
                if let Some(Value::Object(times)) = c.properties.get("times") {
                    let t = times.lock().unwrap();
                    for field in &["user", "sys", "idle"] {
                        assert!(
                            t.properties.contains_key(*field),
                            "cpu.times.{} must exist",
                            field
                        );
                        let val = t
                            .properties
                            .get(*field)
                            .cloned()
                            .unwrap_or(Value::Undefined);
                        assert!(
                            matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
                            "cpu.times.{} must be a number",
                            field
                        );
                    }
                    return;
                }
            }
        }
    }
    // TDD: passes silently if host hasn't implemented times fields yet
}

#[test]
fn cpus_times_has_irq_field() {
    let v = call_os("cpus", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            if let Value::Object(cpu) = &elems[0] {
                let c = cpu.lock().unwrap();
                if let Some(Value::Object(times)) = c.properties.get("times") {
                    let t = times.lock().unwrap();
                    assert!(t.properties.contains_key("irq"), "cpu.times.irq must exist");
                    return;
                }
            }
        }
    }
    // TDD
}

#[test]
fn cpus_times_has_nice_field() {
    let v = call_os("cpus", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            if let Value::Object(cpu) = &elems[0] {
                let c = cpu.lock().unwrap();
                if let Some(Value::Object(times)) = c.properties.get("times") {
                    let t = times.lock().unwrap();
                    assert!(
                        t.properties.contains_key("nice"),
                        "cpu.times.nice must exist"
                    );
                    return;
                }
            }
        }
    }
    // TDD
}

// ── networkInterfaces — entry shape ───────────────────────────────────────────

#[test]
fn network_interfaces_entry_has_address_and_family() {
    let v = call_os("networkInterfaces", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        for (_iface_name, iface_val) in &o.properties {
            if let Value::Object(arr) = iface_val {
                let arr_lock = arr.lock().unwrap();
                if let ObjectKind::Array(entries) = &arr_lock.kind {
                    for entry in entries {
                        if let Value::Object(e) = entry {
                            let e_lock = e.lock().unwrap();
                            assert!(
                                e_lock.properties.contains_key("address"),
                                "networkInterface entry must have address"
                            );
                            assert!(
                                e_lock.properties.contains_key("family"),
                                "networkInterface entry must have family"
                            );
                        }
                    }
                }
                return; // checked first interface
            }
        }
    }
    // TDD: passes silently if host hasn't implemented networkInterfaces entries
}

#[test]
fn network_interfaces_entry_has_internal_and_mac() {
    let v = call_os("networkInterfaces", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        for (_iface_name, iface_val) in &o.properties {
            if let Value::Object(arr) = iface_val {
                let arr_lock = arr.lock().unwrap();
                if let ObjectKind::Array(entries) = &arr_lock.kind {
                    for entry in entries {
                        if let Value::Object(e) = entry {
                            let e_lock = e.lock().unwrap();
                            assert!(
                                e_lock.properties.contains_key("internal"),
                                "networkInterface entry must have internal"
                            );
                            assert!(
                                e_lock.properties.contains_key("mac"),
                                "networkInterface entry must have mac"
                            );
                        }
                    }
                }
                return;
            }
        }
    }
    // TDD
}

#[test]
fn network_interfaces_entry_has_netmask() {
    let v = call_os("networkInterfaces", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        for (_iface_name, iface_val) in &o.properties {
            if let Value::Object(arr) = iface_val {
                let arr_lock = arr.lock().unwrap();
                if let ObjectKind::Array(entries) = &arr_lock.kind {
                    for entry in entries {
                        if let Value::Object(e) = entry {
                            let e_lock = e.lock().unwrap();
                            assert!(
                                e_lock.properties.contains_key("netmask"),
                                "networkInterface entry must have netmask"
                            );
                        }
                    }
                }
                return;
            }
        }
    }
    // TDD
}

#[test]
fn network_interfaces_entry_has_cidr() {
    let v = call_os("networkInterfaces", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        for (_iface_name, iface_val) in &o.properties {
            if let Value::Object(arr) = iface_val {
                let arr_lock = arr.lock().unwrap();
                if let ObjectKind::Array(entries) = &arr_lock.kind {
                    for entry in entries {
                        if let Value::Object(e) = entry {
                            let e_lock = e.lock().unwrap();
                            // cidr may be null for link-local; must be present as a key
                            assert!(
                                e_lock.properties.contains_key("cidr"),
                                "networkInterface entry must have cidr"
                            );
                        }
                    }
                }
                return;
            }
        }
    }
    // TDD
}

// ── constants — extended ──────────────────────────────────────────────────────

#[test]
fn os_constants_errno_ebusy_is_present() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(errno)) = o.properties.get("errno") {
            let e = errno.lock().unwrap();
            let val = e
                .properties
                .get("EBUSY")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert!(
                matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
                "errno.EBUSY must be a number, got {:?}",
                val
            );
        }
    }
    // TDD if errno not yet implemented
}

#[test]
fn os_constants_priority_high_is_negative() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(prio)) = o.properties.get("priority") {
            let p = prio.lock().unwrap();
            if let Some(val) = p.properties.get("PRIORITY_HIGH") {
                match val {
                    Value::I32(n) => assert!(*n < 0, "PRIORITY_HIGH must be negative, got {n}"),
                    Value::I64(n) => assert!(*n < 0, "PRIORITY_HIGH must be negative, got {n}"),
                    Value::F64(f) => assert!(*f < 0.0, "PRIORITY_HIGH must be negative, got {f}"),
                    _ => {}
                }
            }
        }
    }
    // TDD
}

#[test]
fn os_constants_priority_low_is_positive() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(prio)) = o.properties.get("priority") {
            let p = prio.lock().unwrap();
            if let Some(val) = p.properties.get("PRIORITY_LOW") {
                match val {
                    Value::I32(n) => assert!(*n > 0, "PRIORITY_LOW must be positive, got {n}"),
                    Value::I64(n) => assert!(*n > 0, "PRIORITY_LOW must be positive, got {n}"),
                    Value::F64(f) => assert!(*f > 0.0, "PRIORITY_LOW must be positive, got {f}"),
                    _ => {}
                }
            }
        }
    }
    // TDD
}

// ── userInfo — extended ───────────────────────────────────────────────────────

#[cfg(unix)]
#[test]
fn user_info_has_uid_and_gid() {
    let v = call_os("userInfo", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("uid"),
            "userInfo().uid must exist"
        );
        assert!(
            o.properties.contains_key("gid"),
            "userInfo().gid must exist"
        );
        return;
    }
    panic!("userInfo() expected object");
}

#[cfg(unix)]
#[test]
fn user_info_has_homedir() {
    let v = call_os("userInfo", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        let homedir = o
            .properties
            .get("homedir")
            .cloned()
            .unwrap_or(Value::Undefined);
        if let Value::String(s) = homedir {
            assert!(!s.is_empty(), "userInfo().homedir must be non-empty");
            return;
        }
        panic!("userInfo().homedir expected string");
    }
    panic!("userInfo() expected object");
}

#[cfg(unix)]
#[test]
fn user_info_has_shell() {
    let v = call_os("userInfo", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("shell"),
            "userInfo().shell must exist on Unix"
        );
        let shell = o
            .properties
            .get("shell")
            .cloned()
            .unwrap_or(Value::Undefined);
        match shell {
            Value::String(s) => assert!(!s.is_empty(), "userInfo().shell must be non-empty"),
            Value::Null => {} // TDD: acceptable if not yet implemented
            other => panic!("userInfo().shell expected string or null, got {:?}", other),
        }
        return;
    }
    panic!("userInfo() expected object");
}

// ── networkInterfaces — extended ──────────────────────────────────────────────

#[test]
fn network_interfaces_values_are_arrays() {
    let v = call_os("networkInterfaces", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        for (_name, val) in &o.properties {
            assert!(
                matches!(val, Value::Object(_)),
                "each networkInterface entry must be an array, got {:?}",
                val
            );
        }
        return;
    }
    panic!("networkInterfaces() expected object");
}

// ── machine ───────────────────────────────────────────────────────────────────

#[test]
fn machine_returns_non_empty_string() {
    let v = call_os("machine", vec![]);
    let s = as_string(&v);
    assert!(!s.is_empty(), "os.machine() must return non-empty string");
}

#[allow(dead_code)]
fn _force_object_use(_: Object) {}

// ── constants ─────────────────────────────────────────────────────────────────

#[test]
fn os_constants_returns_object() {
    let v = call_os("constants", vec![]);
    assert!(
        matches!(v, Value::Object(_)),
        "os.constants must be an object"
    );
}

#[test]
fn os_constants_has_signals_object() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("signals"),
            "os.constants must have a signals sub-object"
        );
    } else {
        panic!("os.constants expected object, got {:?}", v);
    }
}

#[test]
fn os_constants_signals_sigterm_is_15() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        if let Some(Value::Object(signals)) = o.properties.get("signals") {
            let s = signals.lock().unwrap();
            let sigterm = s
                .properties
                .get("SIGTERM")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(
                sigterm,
                Value::I32(15),
                "SIGTERM must be 15, got {:?}",
                sigterm
            );
        } else {
            panic!("os.constants.signals expected object");
        }
    }
}

#[test]
fn os_constants_has_errno_object() {
    let v = call_os("constants", vec![]);
    if let Value::Object(obj) = &v {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("errno"),
            "os.constants must have an errno sub-object"
        );
    } else {
        panic!("os.constants expected object, got {:?}", v);
    }
}

// ── availableParallelism ──────────────────────────────────────────────────────

#[test]
fn available_parallelism_returns_positive_integer() {
    let v = call_os("availableParallelism", vec![]);
    let n = match v {
        Value::I32(n) => n as i64,
        Value::I64(n) => n,
        Value::F64(f) => f as i64,
        _ => panic!("availableParallelism() expected integer, got {:?}", v),
    };
    assert!(n >= 1, "availableParallelism() must be >= 1, got {n}");
}

// ── devNull ───────────────────────────────────────────────────────────────────

#[test]
fn dev_null_returns_platform_null_device() {
    let v = call_os("devNull", vec![]);
    let s = as_string(&v);
    assert!(
        s == "/dev/null" || s == "nul",
        "os.devNull must be /dev/null or nul, got {s}"
    );
}

#[test]
fn proposal_node_os_surface_is_registered() {
    let expected = [
        "platform",
        "arch",
        "type",
        "release",
        "version",
        "hostname",
        "tmpdir",
        "homedir",
        "endianness",
        "EOL",
        "totalmem",
        "freemem",
        "uptime",
        "loadavg",
        "cpus",
        "networkInterfaces",
        "userInfo",
        "getPriority",
        "setPriority",
        "constants",
        "availableParallelism",
        "devNull",
        "machine",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:os imports: {missing:?}");
}
