/// globalThis, global this binding, module vs script context
use super::helpers::run_js;

#[test]
fn globalthis_is_object() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof globalThis);
"#
        ),
        vec!["object"]
    );
}

#[test]
fn globalthis_holds_globals() {
    assert_eq!(
        run_js(
            r#"
globalThis.myGlobal = 42;
console.log(myGlobal);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn globalthis_consistent_across_accesses() {
    assert_eq!(
        run_js(
            r#"
console.log(globalThis === globalThis);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn undefined_is_a_global() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof undefined);
console.log(undefined === void 0);
"#
        ),
        vec!["undefined", "true"]
    );
}

#[test]
fn nan_is_global_not_a_number() {
    assert_eq!(
        run_js(
            r#"
console.log(isNaN(NaN));
console.log(typeof NaN);
"#
        ),
        vec!["true", "number"]
    );
}

#[test]
fn infinity_is_global_constant() {
    assert_eq!(
        run_js(
            r#"
console.log(Infinity > 1e308);
console.log(-Infinity < -1e308);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn global_isfinite_coerces_argument() {
    assert_eq!(
        run_js(
            r#"
console.log(isFinite("5"));
console.log(isFinite(null));
console.log(isFinite(Infinity));
console.log(isFinite(NaN));
"#
        ),
        vec!["true", "true", "false", "false"]
    );
}

#[test]
fn global_isnan_coerces_argument() {
    assert_eq!(
        run_js(
            r#"
console.log(isNaN("hello"));
console.log(isNaN("5"));
console.log(isNaN(undefined));
"#
        ),
        vec!["true", "false", "true"]
    );
}

#[test]
fn var_in_global_scope_is_on_globalthis() {
    assert_eq!(
        run_js(
            r#"
var declared = "yes";
console.log(globalThis.declared);
"#
        ),
        vec!["yes"]
    );
}

#[test]
fn global_eval_is_function() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof eval);
console.log(typeof Function);
"#
        ),
        vec!["function", "function"]
    );
}

#[test]
fn global_parse_functions_exist() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof parseInt);
console.log(typeof parseFloat);
"#
        ),
        vec!["function", "function"]
    );
}

#[test]
fn global_encode_decode_uri_components() {
    assert_eq!(
        run_js(
            r#"
const enc = encodeURIComponent("hello world");
console.log(enc);
console.log(decodeURIComponent(enc));
"#
        ),
        vec!["hello%20world", "hello world"]
    );
}

#[test]
fn globalthis_undefined_is_read_only() {
    assert_eq!(
        run_js(
            r#"
const desc = Object.getOwnPropertyDescriptor(globalThis, "undefined");
console.log(desc.writable);
"#
        ),
        vec!["false"]
    );
}

