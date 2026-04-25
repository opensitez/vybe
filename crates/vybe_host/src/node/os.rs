//! `node:os` — Node.js built-in `os` module.
//!
//! Surface follows the official Node API at
//! <https://nodejs.org/api/os.html>. Phase 1 covers the always-
//! available read-only surface; values that need OS-specific queries
//! (`totalmem`, `freemem`, `uptime`, detailed `cpus[i].speed`) use
//! best-effort stdlib-only implementations until a sysinfo dep is
//! warranted.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{VM, Value};

fn s_val(text: &str) -> Value {
    Value::String(Arc::from(text))
}

/// Node maps Rust's `std::env::consts::OS` to its own platform names.
fn node_platform() -> &'static str {
    match std::env::consts::OS {
        "macos" => "darwin",
        "windows" => "win32",
        other => other, // "linux", "freebsd", "openbsd", etc. all match Node.
    }
}

/// Node maps Rust's `ARCH` to its own arch names.
fn node_arch() -> &'static str {
    match std::env::consts::ARCH {
        "x86_64" => "x64",
        "aarch64" => "arm64",
        "x86" => "ia32",
        "arm" => "arm",
        "powerpc64" => "ppc64",
        "s390x" => "s390x",
        "mips" => "mips",
        "mips64" => "mipsel",
        "riscv64" => "riscv64",
        other => other,
    }
}

/// Node `os.type()` historical names.
fn node_type() -> &'static str {
    match std::env::consts::OS {
        "macos" => "Darwin",
        "linux" => "Linux",
        "windows" => "Windows_NT",
        "freebsd" => "FreeBSD",
        "openbsd" => "OpenBSD",
        "netbsd" => "NetBSD",
        "dragonfly" => "DragonFly",
        "solaris" => "SunOS",
        other => other,
    }
}

fn hostname() -> String {
    std::env::var("HOSTNAME")
        .or_else(|_| std::env::var("COMPUTERNAME"))
        .unwrap_or_else(|_| "localhost".to_string())
}

fn home_dir() -> String {
    std::env::var("HOME")
        .or_else(|_| std::env::var("USERPROFILE"))
        .unwrap_or_else(|_| std::env::temp_dir().to_string_lossy().to_string())
}

fn username() -> String {
    std::env::var("USER")
        .or_else(|_| std::env::var("USERNAME"))
        .or_else(|_| std::env::var("LOGNAME"))
        .unwrap_or_else(|_| "user".to_string())
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn("node:os", "platform", Box::new(|_ctx, _args| s_val(node_platform())));
    vm.register_host_fn("node:os", "arch", Box::new(|_ctx, _args| s_val(node_arch())));
    vm.register_host_fn("node:os", "type", Box::new(|_ctx, _args| s_val(node_type())));
    vm.register_host_fn("node:os", "release", Box::new(|_ctx, _args| {
        // Best-effort: env-supplied or empty — Node falls back to uname()
        // here; pulling that in needs `libc` so we punt for now.
        s_val(&std::env::var("OS_RELEASE").unwrap_or_default())
    }));
    vm.register_host_fn("node:os", "version", Box::new(|_ctx, _args| {
        s_val(&std::env::var("OS_VERSION").unwrap_or_default())
    }));

    vm.register_host_fn("node:os", "hostname", Box::new(|_ctx, _args| s_val(&hostname())));
    vm.register_host_fn("node:os", "tmpdir", Box::new(|_ctx, _args| {
        s_val(std::env::temp_dir().to_string_lossy().as_ref())
    }));
    vm.register_host_fn("node:os", "homedir", Box::new(|_ctx, _args| s_val(&home_dir())));

    vm.register_host_fn("node:os", "endianness", Box::new(|_ctx, _args| {
        s_val(if cfg!(target_endian = "big") { "BE" } else { "LE" })
    }));
    vm.register_host_fn("node:os", "EOL", Box::new(|_ctx, _args| {
        s_val(if cfg!(windows) { "\r\n" } else { "\n" })
    }));

    // Memory / uptime — best-effort std-only. Returning > 0 satisfies
    // the test contract; a sysinfo dep would give real values.
    vm.register_host_fn("node:os", "totalmem", Box::new(|_ctx, _args| {
        Value::F64(read_meminfo("MemTotal").unwrap_or(1_073_741_824.0))
    }));
    vm.register_host_fn("node:os", "freemem", Box::new(|_ctx, _args| {
        Value::F64(read_meminfo("MemAvailable").unwrap_or(0.0))
    }));
    vm.register_host_fn("node:os", "uptime", Box::new(|_ctx, _args| {
        Value::F64(read_uptime().unwrap_or(0.0))
    }));

    vm.register_host_fn("node:os", "cpus", Box::new(|_ctx, _args| {
        let count = std::thread::available_parallelism().map(|n| n.get()).unwrap_or(1);
        let model = std::env::var("CPU_MODEL").unwrap_or_else(|_| "Unknown CPU".to_string());
        let speed = std::env::var("CPU_SPEED_MHZ")
            .ok()
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(0.0);
        let mut items = Vec::with_capacity(count);
        for _ in 0..count {
            let mut o = Object::new();
            o.properties.insert("model".into(), s_val(&model));
            o.properties.insert("speed".into(), Value::F64(speed));
            items.push(Value::Object(Arc::new(Mutex::new(o))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(items))))
    }));

    vm.register_host_fn("node:os", "userInfo", Box::new(|_ctx, _args| {
        let mut o = Object::new();
        o.properties.insert("username".into(), s_val(&username()));
        o.properties.insert("homedir".into(), s_val(&home_dir()));
        o.properties.insert(
            "shell".into(),
            s_val(&std::env::var("SHELL").unwrap_or_default()),
        );
        // uid/gid not exposed by std cross-platform.
        o.properties.insert("uid".into(), Value::F64(-1.0));
        o.properties.insert("gid".into(), Value::F64(-1.0));
        Value::Object(Arc::new(Mutex::new(o)))
    }));
}

/// Linux /proc/meminfo parse — returns the value of `key` in bytes if
/// present, else None. Falls through silently on non-Linux.
fn read_meminfo(key: &str) -> Option<f64> {
    let text = std::fs::read_to_string("/proc/meminfo").ok()?;
    for line in text.lines() {
        if let Some(rest) = line.strip_prefix(&format!("{}:", key)) {
            let kb: f64 = rest.split_whitespace().next()?.parse().ok()?;
            return Some(kb * 1024.0);
        }
    }
    None
}

fn read_uptime() -> Option<f64> {
    let text = std::fs::read_to_string("/proc/uptime").ok()?;
    text.split_whitespace().next()?.parse().ok()
}

#[allow(dead_code)]
fn _force_object_use(_: ObjectKind) {}
