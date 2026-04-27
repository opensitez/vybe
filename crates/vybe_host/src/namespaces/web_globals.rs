//! Web platform globals — wires up the `web:*` host modules as JS-style
//! global objects and constructors:
//!
//! ```js
//! crypto.randomUUID();
//! const url = new URL("https://example.com/path");
//! const enc = new TextEncoder();
//! const res = await fetch("https://api.example.com/x");
//! const id  = setTimeout(fn, 100);
//! ```
//!
//! Every JS runtime (Node, Deno, Bun, browsers) exposes these as
//! ambient globals — Vybe matches by registering them on the VM
//! globals table here.

use super::*;

pub fn register(vm: &mut VM) {
    // ── crypto ─────────────────────────────────────────────────────
    let crypto = ensure_namespace(vm, &["crypto"]);
    set_prop(&crypto, "randomUUID",      host_fn_ref(vm, "web:crypto", "randomUUID"));
    set_prop(&crypto, "getRandomValues", host_fn_ref(vm, "web:crypto", "getRandomValues"));

    // crypto.subtle.digest — nested namespace per WebCryptoAPI §13.
    let subtle = ensure_namespace(vm, &["crypto", "subtle"]);
    set_prop(&subtle, "digest", host_fn_ref(vm, "web:crypto", "digest"));

    // ── URL + URLSearchParams ──────────────────────────────────────
    // Vybe's `new URL(...)` known_types entry routes to web:url.new
    // (see js/profile). Static methods + the URL global itself live here.
    let url = ensure_namespace(vm, &["URL"]);
    set_prop(&url, "parse",    host_fn_ref(vm, "web:url", "parse"));
    set_prop(&url, "canParse", host_fn_ref(vm, "web:url", "canParse"));

    // TextEncoder / TextDecoder — empty namespaces; constructed via
    // known_types entries that point at web:encoding.{encoderNew, decoderNew}.
    // The empty namespace is still useful so `typeof TextEncoder === "function"`
    // holds (resolved as object presence in Vybe's typeof).
    let _ = ensure_namespace(vm, &["TextEncoder"]);
    let _ = ensure_namespace(vm, &["TextDecoder"]);
    let _ = ensure_namespace(vm, &["URLSearchParams"]);

    // ── fetch + Headers/Request/Response ───────────────────────────
    // `fetch(...)` is a top-level function, not a namespace.
    vm.globals.insert("fetch".to_string(), host_fn_ref(vm, "web:fetch", "fetch"));

    let _ = ensure_namespace(vm, &["Headers"]);
    let _ = ensure_namespace(vm, &["Request"]);
    let _ = ensure_namespace(vm, &["Response"]);

    // ── Timers ─────────────────────────────────────────────────────
    vm.globals.insert("setTimeout".to_string(),    host_fn_ref(vm, "web:timers", "setTimeout"));
    vm.globals.insert("clearTimeout".to_string(),  host_fn_ref(vm, "web:timers", "clearTimeout"));
    vm.globals.insert("setInterval".to_string(),   host_fn_ref(vm, "web:timers", "setInterval"));
    vm.globals.insert("clearInterval".to_string(), host_fn_ref(vm, "web:timers", "clearInterval"));
    vm.globals.insert("queueMicrotask".to_string(), host_fn_ref(vm, "web:timers", "queueMicrotask"));
}
