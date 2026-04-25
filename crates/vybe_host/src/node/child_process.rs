//! `node:child_process` — Node.js built-in `child_process` module.
//!
//! Reference: <https://nodejs.org/api/child_process.html>.
//!
//! Phase 1 covers sync variants only: `spawnSync`, `execSync`,
//! `execFileSync`. Async (callback) and stream (`spawn`/`exec` w/
//! pipes) variants come later.

use std::process::Command;
use std::sync::{Arc, Mutex};
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

fn extract_string_array(value: &Value) -> Vec<String> {
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

fn opt_string_property(args: &[Value], opt_idx: usize, key: &str) -> Option<String> {
    if let Some(Value::Object(obj)) = args.get(opt_idx) {
        let o = obj.lock().unwrap();
        if let Some(Value::String(text)) = o.properties.get(key) {
            return Some(text.to_string());
        }
    }
    None
}

pub fn register(vm: &mut VM) {
    // ── execSync(command[, options]) ──────────────────────────────
    // Runs `command` via the platform shell. Returns stdout as a
    // string (when `options.encoding` is set) or as a byte-array
    // surrogate for Buffer otherwise.
    vm.register_host_fn("node:child_process", "execSync", Box::new(|_ctx, args| {
        let command = s_arg(args, 0, "");
        let encoding = opt_string_property(args, 1, "encoding");

        let output = if cfg!(windows) {
            Command::new("cmd").args(["/d", "/s", "/c", &command]).output()
        } else {
            Command::new("sh").args(["-c", &command]).output()
        };

        match output {
            Ok(o) => match encoding.as_deref() {
                Some("utf8") | Some("utf-8") | Some("UTF-8") => {
                    s_val(&String::from_utf8_lossy(&o.stdout))
                }
                None => {
                    // Default: Node returns a Buffer; we use a string
                    // here too because most consumers .toString() on it
                    // anyway. Spec-strict Buffer return is deferred
                    // until a Buffer type lands.
                    s_val(&String::from_utf8_lossy(&o.stdout))
                }
                _ => s_val(&String::from_utf8_lossy(&o.stdout)),
            },
            Err(e) => s_val(&format!("execSync error: {}", e)),
        }
    }));

    // ── execFileSync(file[, args[, options]]) ─────────────────────
    // Runs `file` directly (no shell) with optional args. Same return
    // contract as execSync.
    vm.register_host_fn("node:child_process", "execFileSync", Box::new(|_ctx, args| {
        let file = s_arg(args, 0, "");
        let cmd_args: Vec<String> = match args.get(1) {
            Some(arr) => extract_string_array(arr),
            None => Vec::new(),
        };
        let encoding = opt_string_property(args, 2, "encoding");

        let mut cmd = Command::new(&file);
        for a in &cmd_args { cmd.arg(a); }

        match cmd.output() {
            Ok(o) => match encoding.as_deref() {
                Some("utf8") | Some("utf-8") | Some("UTF-8") => {
                    s_val(&String::from_utf8_lossy(&o.stdout))
                }
                _ => s_val(&String::from_utf8_lossy(&o.stdout)),
            },
            Err(e) => s_val(&format!("execFileSync error: {}", e)),
        }
    }));

    // ── spawnSync(command[, args[, options]]) ─────────────────────
    // Returns:
    //   { pid, output, stdout, stderr, status, signal, error }
    vm.register_host_fn("node:child_process", "spawnSync", Box::new(|_ctx, args| {
        let command = s_arg(args, 0, "");
        let cmd_args: Vec<String> = match args.get(1) {
            Some(arr) => extract_string_array(arr),
            None => Vec::new(),
        };

        let mut cmd = Command::new(&command);
        for a in &cmd_args { cmd.arg(a); }

        let mut o = Object::new();
        match cmd.output() {
            Ok(out) => {
                let stdout = String::from_utf8_lossy(&out.stdout).to_string();
                let stderr = String::from_utf8_lossy(&out.stderr).to_string();
                let status = out.status.code().unwrap_or(-1);
                // pid is consumed by .output(); return the current process
                // PID + 1 as a placeholder until a streaming spawn is
                // implemented (Node returns the actual child PID).
                o.properties.insert("pid".into(), Value::F64((std::process::id() as f64) + 1.0));
                o.properties.insert("stdout".into(), s_val(&stdout));
                o.properties.insert("stderr".into(), s_val(&stderr));
                o.properties.insert("status".into(), Value::F64(status as f64));
                o.properties.insert("signal".into(), Value::Null);
                o.properties.insert("error".into(), Value::Null);
                let output_arr = vec![
                    Value::Null,        // index 0 = stdin (always null in Node)
                    s_val(&stdout),
                    s_val(&stderr),
                ];
                o.properties.insert(
                    "output".into(),
                    Value::Object(Arc::new(Mutex::new(Object::new_array(output_arr)))),
                );
            }
            Err(e) => {
                o.properties.insert("pid".into(), Value::F64(0.0));
                o.properties.insert("stdout".into(), s_val(""));
                o.properties.insert("stderr".into(), s_val(""));
                o.properties.insert("status".into(), Value::Null);
                o.properties.insert("signal".into(), Value::Null);
                o.properties.insert("error".into(), s_val(&format!("{}", e)));
            }
        }
        Value::Object(Arc::new(Mutex::new(o)))
    }));
}
