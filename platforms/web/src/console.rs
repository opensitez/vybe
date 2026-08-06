//! `web:console` — the WHATWG Console Standard surface (§1.1 Logging
//! functions). `log(...data)` is VARIADIC BY SPEC — each datum is rendered
//! and the results are joined with a single space. This is the home for
//! every console/print-shaped lowering (`console.log`, Python `print`,
//! Ruby `puts`, PHP `echo`, …).
//!
//! Rendering is the VM's console/inspect surface (`Value`'s `Display`:
//! BigInt `8n`, `-0`, arrays `1,2`, …) — NOT ECMAScript `ToString`; the
//! rendering moved here VERBATIM from the old `wasi:logging` heuristic so
//! every existing output expectation is byte-identical.
//!
//! `wasi:logging/logging.log` is the STRICT spec interface
//! `log(level, context, message)` — it no longer arity-sniffs, which is
//! what silently dropped arguments from 2- and 3-arg console calls
//! (`console.log("a", "b")` printed only `b`).

use vybe_runtime::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    // log(...data) — WHATWG §1.1.6: Logger("log", data). Space-joined per
    // the Printer's inspection of each datum.
    vm.register_host_fn(
        "web:console",
        "log",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
            println!("{}", parts.join(" "));
            Value::Null
        }),
    );

    // error(...data) — WHATWG §1.1.2: same rendering, stderr stream.
    vm.register_host_fn(
        "web:console",
        "error",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
            eprintln!("{}", parts.join(" "));
            Value::Null
        }),
    );
}
