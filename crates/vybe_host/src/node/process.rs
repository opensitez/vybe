//! `node:process` — Node.js built-in `process` module.
//!
//! Reference: <https://nodejs.org/api/process.html>.
//!
//! Phase 1 covers always-available reads + a few mutators. Streams,
//! event listeners, and `nextTick` are deferred.

use std::sync::{Arc, Mutex, OnceLock};
use std::time::Instant;
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{VM, Value};

fn s_arg(args: &[Value], idx: usize, default: &str) -> String {
    match args.get(idx) {
        Some(Value::String(text)) => text.to_string(),
        Some(other) => format!("{}", other),
        None => default.to_string(),
    }
}

fn s_val(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn versions_value() -> Value {
    let mut object = Object::new();
    object.properties.insert("vybe".into(), s_val(env!("CARGO_PKG_VERSION")));
    object.properties.insert("node".into(), s_val(env!("CARGO_PKG_VERSION")));
    Value::Object(Arc::new(Mutex::new(object)))
}

fn argv_value() -> Value {
    let items: Vec<Value> = std::env::args()
        .map(|arg| s_val(&arg))
        .collect();
    Value::Object(Arc::new(Mutex::new(Object::new_array(items))))
}

fn argv0_value() -> Value {
    let argv0 = std::env::args().next().unwrap_or_else(|| "vybex".to_string());
    s_val(&argv0)
}

fn exec_path_value() -> Value {
    match std::env::current_exe() {
        Ok(path) => s_val(path.to_string_lossy().as_ref()),
        Err(_) => s_val(""),
    }
}

fn env_value() -> Value {
    let mut object = Object::new();
    for (key, value) in std::env::vars() {
        object.properties.insert(key, s_val(&value));
    }
    Value::Object(Arc::new(Mutex::new(object)))
}

/// Process start time, captured once at module init so `uptime()`
/// can return monotonic seconds since the host process began.
fn process_start() -> Instant {
    static START: OnceLock<Instant> = OnceLock::new();
    *START.get_or_init(Instant::now)
}

fn node_platform() -> &'static str {
    match std::env::consts::OS {
        "macos" => "darwin",
        "windows" => "win32",
        other => other,
    }
}

fn node_arch() -> &'static str {
    match std::env::consts::ARCH {
        "x86_64" => "x64",
        "aarch64" => "arm64",
        "x86" => "ia32",
        "arm" => "arm",
        "powerpc64" => "ppc64",
        "s390x" => "s390x",
        "riscv64" => "riscv64",
        other => other,
    }
}

pub fn register(vm: &mut VM) {
    let _ = process_start(); // initialize the start-time cache up front

    vm.register_host_value("node:process", "platform", s_val(node_platform()));
    vm.register_host_value("node:process", "arch", s_val(node_arch()));

    // Vybe doesn't ship as Node, but consumers expect a "v"-prefixed
    // version string for compat (most regex parsing assumes `v(\d+)`).
    vm.register_host_value("node:process", "version", s_val(concat!("v", env!("CARGO_PKG_VERSION"))));
    vm.register_host_value("node:process", "versions", versions_value());
    vm.register_host_value("node:process", "pid", Value::F64(std::process::id() as f64));
    // Cross-platform parent PID requires libc/winapi calls. Punt until a
    // test actually demands it; return 0 for now (Node returns the real PPID).
    vm.register_host_value("node:process", "ppid", Value::F64(0.0));

    vm.register_host_fn("node:process", "cwd", Box::new(|_ctx, _args| {
        match std::env::current_dir() {
            Ok(p) => s_val(p.to_string_lossy().as_ref()),
            Err(_) => s_val("."),
        }
    }));

    vm.register_host_fn("node:process", "chdir", Box::new(|_ctx, args| {
        let path = s_arg(args, 0, ".");
        let _ = std::env::set_current_dir(&path);
        Value::Null
    }));

    vm.register_host_fn("node:process", "exit", Box::new(|_ctx, args| {
        let code = match args.first() {
            Some(Value::F64(n)) => *n as i32,
            Some(Value::I32(n)) => *n,
            _ => 0,
        };
        std::process::exit(code);
    }));

    vm.register_host_value("node:process", "argv", argv_value());
    vm.register_host_value("node:process", "argv0", argv0_value());
    vm.register_host_value("node:process", "execPath", exec_path_value());
    vm.register_host_value("node:process", "env", env_value());

    vm.register_host_fn("node:process", "uptime", Box::new(|_ctx, _args| {
        Value::F64(process_start().elapsed().as_secs_f64())
    }));

    vm.register_host_fn("node:process", "hrtime", Box::new(|_ctx, _args| {
        // Node hrtime() returns [seconds, nanoseconds] since process start.
        let elapsed = process_start().elapsed();
        let pair = vec![
            Value::F64(elapsed.as_secs() as f64),
            Value::F64(elapsed.subsec_nanos() as f64),
        ];
        Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
    }));

    vm.register_host_fn("node:process", "memoryUsage", Box::new(|_ctx, _args| {
        // Node returns {rss, heapTotal, heapUsed, external, arrayBuffers}.
        // Vybe doesn't have a JS heap; report zeros for fields we can't
        // measure cheaply. RSS could be read from /proc/self/status on
        // Linux (best-effort).
        let rss = read_linux_rss().unwrap_or(0.0);
        let mut o = Object::new();
        o.properties.insert("rss".into(), Value::F64(rss));
        o.properties.insert("heapTotal".into(), Value::F64(0.0));
        o.properties.insert("heapUsed".into(), Value::F64(0.0));
        o.properties.insert("external".into(), Value::F64(0.0));
        o.properties.insert("arrayBuffers".into(), Value::F64(0.0));
        Value::Object(Arc::new(Mutex::new(o)))
    }));
}

/// `/proc/self/status` `VmRSS:` parse — bytes resident set size.
/// Returns None on non-Linux or when the field is absent.
fn read_linux_rss() -> Option<f64> {
    let text = std::fs::read_to_string("/proc/self/status").ok()?;
    for line in text.lines() {
        if let Some(rest) = line.strip_prefix("VmRSS:") {
            let kb: f64 = rest.split_whitespace().next()?.parse().ok()?;
            return Some(kb * 1024.0);
        }
    }
    None
}

#[allow(dead_code)]
fn _force_object_use(_: ObjectKind) {}
