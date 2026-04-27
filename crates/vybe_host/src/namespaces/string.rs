use super::*;

pub fn register(vm: &mut VM) {
    // Strings (Microsoft.VisualBasic.Strings shortcut)
    let strings = ensure_namespace(vm, &["Strings"]);
    for (vb_name, host_name) in &[
        ("left", "left"), ("right", "right"), ("mid", "mid"),
        ("instr", "instr"), ("ucase", "ucase"), ("lcase", "lcase"),
        ("trim", "trim"), ("ltrim", "ltrim"), ("rtrim", "rtrim"),
        ("len", "length"), ("asc", "asc"), ("chr", "chr"),
        ("space", "space"), ("replace", "replaceAll"), ("split", "split"),
    ] {
        set_prop(&strings, vb_name, host_fn_ref(vm, "vybe:string", host_name));
    }

    // System.String
    let sys = ensure_namespace(vm, &["System", "String"]);
    set_prop(&sys, "isnullorempty", host_fn_ref(vm, "vybe:string", "length"));
    set_prop(&sys, "format", host_fn_ref(vm, "vybe:string", "format"));
    set_prop(&sys, "join", host_fn_ref(vm, "ecma:array", "join"));

    // Encoding (System.Text.Encoding)
    let utf8 = ensure_namespace(vm, &["Encoding", "UTF8"]);
    set_prop(&utf8, "getbytes", host_fn_ref(vm, "vybe:convert", "toString"));
    set_prop(&utf8, "getstring", host_fn_ref(vm, "vybe:convert", "toString"));

    let ascii = ensure_namespace(vm, &["Encoding", "ASCII"]);
    set_prop(&ascii, "getbytes", host_fn_ref(vm, "vybe:convert", "toString"));
    set_prop(&ascii, "getstring", host_fn_ref(vm, "vybe:convert", "toString"));

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
