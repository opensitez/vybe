use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Symbol.for()` & `Symbol.keyFor()` Global Registry
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_symbol_for_global_registry_identity() {
    let src = r#"
const s1 = Symbol.for("app.key");
const s2 = Symbol.for("app.key");
console.log(s1 === s2);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_keyfor_returns_registered_string() {
    let src = r#"
const sym = Symbol.for("user.token");
console.log(Symbol.keyFor(sym));
"#;
    assert_eq!(run_js(src), vec!["user.token"]);
}

#[test]
fn test_js_symbol_keyfor_local_symbol_returns_undefined() {
    let src = r#"
const localSym = Symbol("user.token");
console.log(Symbol.keyFor(localSym) === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_keyfor_non_symbol_throws_typeerror() {
    let src = r#"
try {
    Symbol.keyFor("not_a_symbol");
} catch (e) {
    console.log("Symbol.keyFor Non-Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Symbol.keyFor Non-Symbol TypeError"]);
}

#[test]
fn test_js_symbol_for_coerces_key_argument_to_string() {
    let src = r#"
const s1 = Symbol.for(100);
const s2 = Symbol.for("100");
console.log(s1 === s2 + "|" + Symbol.keyFor(s1));
"#;
    assert_eq!(run_js(src), vec!["true|100"]);
}

#[test]
fn test_js_symbol_description_property() {
    let src = r#"
const s1 = Symbol("desc1");
const s2 = Symbol.for("desc2");
const s3 = Symbol();
console.log(`${s1.description}:${s2.description}:${s3.description}`);
"#;
    assert_eq!(run_js(src), vec!["desc1:desc2:undefined"]);
}

#[test]
fn test_js_symbol_description_read_only() {
    let src = r#"
const sym = Symbol("orig");
try {
    "use strict";
    sym.description = "new";
} catch (e) {
    console.log("Description Read-Only TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Description Read-Only TypeError"]);
}

#[test]
fn test_js_symbol_cannot_be_constructed_with_new() {
    let src = r#"
try {
    eval("new Symbol('bad')");
} catch (e) {
    console.log("New Symbol TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["New Symbol TypeError"]);
}

#[test]
fn test_js_symbol_typeof_operator() {
    let src = r#"
const s = Symbol();
console.log(typeof s);
"#;
    assert_eq!(run_js(src), vec!["symbol"]);
}

#[test]
fn test_js_symbol_wrapper_object_explicit_coercion() {
    let src = r#"
const s = Symbol("wrapped");
const obj = Object(s);
console.log(typeof obj + "|" + (obj.valueOf() === s));
"#;
    assert_eq!(run_js(src), vec!["object|true"]);
}

#[test]
fn test_js_symbol_property_key_in_object_literal() {
    let src = r#"
const sym = Symbol.for("id");
const obj = { [sym]: 42 };
console.log(obj[sym] + "|" + obj[Symbol.for("id")]);
"#;
    assert_eq!(run_js(src), vec!["42|42"]);
}

#[test]
fn test_js_symbol_getownpropertysymbols_utility() {
    let src = r#"
const s1 = Symbol("a");
const s2 = Symbol.for("b");
const obj = { [s1]: 1, [s2]: 2, stringKey: 3 };

const symbols = Object.getOwnPropertySymbols(obj);
console.log(symbols.length + "|" + (symbols[0] === s1) + "|" + (symbols[1] === s2));
"#;
    assert_eq!(run_js(src), vec!["2|true|true"]);
}

#[test]
fn test_js_symbol_keys_ignored_by_for_in_and_object_keys() {
    let src = r#"
const sym = Symbol("hidden");
const obj = { [sym]: 10, pub: 20 };
console.log(Object.keys(obj).join(",") + "|hasIn=" + (sym in obj));
"#;
    assert_eq!(run_js(src), vec!["pub|hasIn=true"]);
}

#[test]
fn test_js_symbol_json_stringify_omits_symbol_properties() {
    let src = r#"
const sym = Symbol("secret");
const obj = { [sym]: "hidden", publicData: "visible" };
console.log(JSON.stringify(obj));
"#;
    assert_eq!(run_js(src), vec![r#"{"publicData":"visible"}"#]);
}

#[test]
fn test_js_symbol_unique_identity_per_call() {
    let src = r#"
const s1 = Symbol("same");
const s2 = Symbol("same");
console.log(s1 === s2);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_symbol_tostring_explicit_conversion() {
    let src = r#"
const sym = Symbol("tag");
console.log(sym.toString() + "|" + String(sym));
"#;
    assert_eq!(run_js(src), vec!["Symbol(tag)|Symbol(tag)"]);
}

#[test]
fn test_js_symbol_implicit_string_concatenation_throws_typeerror() {
    let src = r#"
const sym = Symbol("err");
try {
    const msg = "Symbol: " + sym;
} catch (e) {
    console.log("Implicit Symbol String Coercion TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Implicit Symbol String Coercion TypeError"]
    );
}

#[test]
fn test_js_symbol_to_boolean_coercion_always_true() {
    let src = r#"
const s1 = Symbol("");
const s2 = Symbol.for("test");
console.log(Boolean(s1) + "|" + Boolean(s2));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_symbol_to_number_coercion_throws_typeerror() {
    let src = r#"
const sym = Symbol("num");
try {
    const n = Number(sym);
} catch (e) {
    console.log("Symbol to Number TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Symbol to Number TypeError"]);
}

#[test]
fn test_js_symbol_reflect_ownkeys_includes_symbols() {
    let src = r#"
const s = Symbol("s");
const obj = { a: 1, [s]: 2 };
const keys = Reflect.ownKeys(obj);
console.log(keys.length + "|" + (keys[1] === s));
"#;
    assert_eq!(run_js(src), vec!["2|true"]);
}

#[test]
fn test_js_symbol_keyfor_empty_string_key() {
    let src = r#"
const s = Symbol.for("");
console.log(Symbol.keyFor(s) === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

