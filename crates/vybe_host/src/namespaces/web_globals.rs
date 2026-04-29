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

    // ── DOMParser / XMLSerializer (WHATWG DOM Parsing & Serialization) ──
    // `new DOMParser()` ctor lowers via `known_types` to
    // `web:dom-parser.parserNew`; static-like helpers live on the
    // namespace object so `DOMParser.parseFromString(s)` works without
    // an instance for languages that prefer flat dispatch.
    let dom_parser = ensure_namespace(vm, &["DOMParser"]);
    set_prop(&dom_parser, "parseFromString", host_fn_ref(vm, "web:dom-parser", "parseFromString"));
    let xml_serializer = ensure_namespace(vm, &["XMLSerializer"]);
    set_prop(&xml_serializer, "serializeToString", host_fn_ref(vm, "web:dom-parser", "serializeToString"));

    // `xml` shorthand — Vybe-side convenience that mirrors what
    // `import { parse } from "xml"` resolved to historically. Same fns
    // as `web:dom-parser/*` so language profile entries that reference
    // `xml.parse` / `xml.parseString` resolve at runtime even before
    // dotted-name canonicalisation kicks in.
    let xml = ensure_namespace(vm, &["xml"]);
    set_prop(&xml, "parse",       host_fn_ref(vm, "web:dom-parser", "parse"));
    set_prop(&xml, "parseString", host_fn_ref(vm, "web:dom-parser", "parse"));
    set_prop(&xml, "load",        host_fn_ref(vm, "web:dom-parser", "load"));
    set_prop(&xml, "toString",    host_fn_ref(vm, "web:dom-parser", "toString"));

    // .NET-style aliases — VB / C# tests use `XDocument.Parse(s)` /
    // `XmlDocument.LoadXml(s)`. The dotnet wrapper handles the typed
    // case at compile time via component_classes, but the namespace
    // objects themselves provide static-method dispatch for untyped
    // identifier resolution.
    let xdocument = ensure_namespace(vm, &["XDocument"]);
    set_prop(&xdocument, "Parse",    host_fn_ref(vm, "web:dom-parser", "parse"));
    set_prop(&xdocument, "Load",     host_fn_ref(vm, "web:dom-parser", "load"));

    let xml_document = ensure_namespace(vm, &["XmlDocument"]);
    set_prop(&xml_document, "LoadXml", host_fn_ref(vm, "web:dom-parser", "parse"));
    set_prop(&xml_document, "Load",    host_fn_ref(vm, "web:dom-parser", "load"));
}
