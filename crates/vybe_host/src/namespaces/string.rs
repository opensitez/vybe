use super::*;

pub fn register(vm: &mut VM) {
    // Microsoft.VisualBasic.Strings — bound to ECMA-262 §22.1
    // String.prototype methods. VB-shape divergence (1-based indices in
    // Mid/InStr, Left/Right slice anchoring) is a language-adapter
    // concern at compile time, not host-fn business.
    let strings = ensure_namespace(vm, &["Strings"]);
    for (vb_name, ecma_name) in &[
        ("left", "slice"),          // VB Left(s, n) ≈ s.slice(0, n) at adapter
        ("right", "slice"),         // VB Right(s, n) ≈ s.slice(-n) at adapter
        ("mid", "substring"),       // VB Mid(s, start[, len]) — 1-based at adapter
        ("instr", "indexOf"),       // VB InStr returns 1-based; +1 at adapter
        ("ucase", "toUpperCase"),
        ("lcase", "toLowerCase"),
        ("trim", "trim"),
        ("ltrim", "trimStart"),
        ("rtrim", "trimEnd"),
        ("len", "length"),
        ("asc", "charCodeAt"),
        ("chr", "fromCharCode"),
        ("space", "repeat"),        // VB Space(n) ≈ " ".repeat(n) at adapter
        ("replace", "replaceAll"),
        ("split", "split"),
    ] {
        set_prop(&strings, vb_name, host_fn_ref(vm, "ecma:string", ecma_name));
    }

    // System.String — `isnullorempty` / `format` have no single ECMA host
    // fn target; bindings dropped (language adapters compile them inline).
    let sys = ensure_namespace(vm, &["System", "String"]);
    set_prop(&sys, "join", host_fn_ref(vm, "ecma:array", "join"));

    // Encoding (System.Text.Encoding) — bound to WHATWG `web:encoding`
    // (TextEncoder / TextDecoder). UTF-8 only at the encoder; the
    // decoder picks up encoding at construction so ASCII shares the
    // same `decode` host fn (callers wanting strict-ASCII validation
    // do it in adapter code).
    let utf8 = ensure_namespace(vm, &["Encoding", "UTF8"]);
    set_prop(&utf8, "getbytes", host_fn_ref(vm, "web:encoding", "encode"));
    set_prop(&utf8, "getstring", host_fn_ref(vm, "web:encoding", "decode"));

    let ascii = ensure_namespace(vm, &["Encoding", "ASCII"]);
    set_prop(&ascii, "getbytes", host_fn_ref(vm, "web:encoding", "encode"));
    set_prop(&ascii, "getstring", host_fn_ref(vm, "web:encoding", "decode"));

    // System.Text.RegularExpressions.Regex — bound to ECMA-262 §22.2 RegExp.
    // .NET callers typically pass `(input, pattern)` (input-first per the
    // .NET shape), but `ecma:regexp.test/exec` accept `(pattern, input)` —
    // the wrong order would yield a regex that matches the literal pattern
    // string instead of the intended regex. Direct binding works for callers
    // that have already adapted to pattern-first; compile-site adapter chunks
    // handle the rest (see emitter/stdlib.rs::build_regex_*_pat_first).
    let regex = ensure_namespace(vm, &["System", "Text", "RegularExpressions", "Regex"]);
    set_prop(&regex, "ismatch", host_fn_ref(vm, "ecma:regexp", "test"));
    set_prop(&regex, "match", host_fn_ref(vm, "ecma:regexp", "exec"));
    set_prop(&regex, "replace", host_fn_ref(vm, "ecma:regexp", "replace"));
    set_prop(&regex, "split", host_fn_ref(vm, "ecma:regexp", "split"));
}
