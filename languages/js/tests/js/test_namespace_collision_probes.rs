//! Namespace-collision probes (namespaceplan.md Phase 1).
//!
//! Each case is a historical resolution hijack: a user binding whose name
//! collides with a host namespace / builtin surface must ALWAYS win
//! (locals shadow, always), and a named function expression's name binds
//! only in its own body — it must never poison module-wide dispatch.
//! These lock in the fixed behavior so the resolver migration can never
//! silently reintroduce the bug class.

crate::js_cases! {
    const_string_local_shadows_host_namespace => {
        r#"const string = "ab,cd"; console.log(string);"#,
        ["ab,cd"]
    };
    const_math_local_shadows_host_namespace => {
        r#"const math = 5; console.log(math + 1);"#,
        ["6"]
    };
    const_text_receiver_method_not_hijacked => {
        r#"const text = "a1b2"; console.log([...text.matchAll(/\d/g)].length);"#,
        ["2"]
    };
    named_fn_expr_tostring_does_not_poison_dispatch => {
        r#"const f = function toString() { return 1; };
console.log(String({ x: 1 }));
console.log(f());"#,
        ["[object Object]", "1"]
    };
    named_fn_expr_name_binds_only_in_own_body => {
        r#"const f = function named() { return 1; };
let out;
try { named(); out = "leaked"; } catch (e) { out = "scoped"; }
console.log(out);"#,
        ["scoped"]
    };
    generator_stored_in_array_keeps_prototype => {
        r#"function* g() { yield 1; yield 2; }
const arr = [g];
console.log([...arr[0]()].length);"#,
        ["2"]
    };
    const_values_local_spread_not_hijacked => {
        r#"const values = [9, 8]; console.log([...values].length);"#,
        ["2"]
    };
    local_console_object_shadows_wasi_alias => {
        r#"{
  let console2 = { log: (x) => x };
  console.log(console2.log("shadow-ok"));
}"#,
        ["shadow-ok"]
    };

    const_array_local_shadows_host_constructor => {
        r#"{
  const Array = 42;
  console.log(Array);
}"#,
        ["42"]
    };
}
