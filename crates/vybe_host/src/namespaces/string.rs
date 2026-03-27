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
    set_prop(&sys, "join", host_fn_ref(vm, "vybe:array", "join"));

    // System.Text.RegularExpressions.Regex
    let regex = ensure_namespace(vm, &["System", "Text", "RegularExpressions", "Regex"]);
    set_prop(&regex, "ismatch", host_fn_ref(vm, "vybe:regex", "test"));
    set_prop(&regex, "match", host_fn_ref(vm, "vybe:regex", "match"));
    set_prop(&regex, "replace", host_fn_ref(vm, "vybe:regex", "replace"));
    set_prop(&regex, "split", host_fn_ref(vm, "vybe:regex", "split"));
}
