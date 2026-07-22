use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `globalThis` Standardized Global Scope Access & Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_globalthis_identity() {
    let src = r#"
console.log(globalThis === globalThis.globalThis);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_globalthis_defines_global_variable() {
    let src = r#"
globalThis.customGlobalVar = 12345;
console.log(customGlobalVar + "|" + globalThis.customGlobalVar);
"#;
    assert_eq!(run_js(src), vec!["12345|12345"]);
}

#[test]
fn test_js_globalthis_var_declarations_attached_to_globalthis() {
    let src = r#"
var explicitVar = "attached";
console.log(globalThis.explicitVar);
"#;
    assert_eq!(run_js(src), vec!["attached"]);
}

#[test]
fn test_js_globalthis_let_and_const_not_attached_to_globalthis() {
    let src = r#"
let lexicalLet = "unattachedLet";
const lexicalConst = "unattachedConst";
console.log((globalThis.lexicalLet === undefined) + "|" + (globalThis.lexicalConst === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_globalthis_builtin_constructors_accessible() {
    let src = r#"
console.log(`${globalThis.Object === Object}:${globalThis.Array === Array}:${globalThis.Function === Function}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]);
}

#[test]
fn test_js_globalthis_typeof_operator() {
    let src = r#"
console.log(typeof globalThis);
"#;
    assert_eq!(run_js(src), vec!["object"]);
}

#[test]
fn test_js_globalthis_function_declaration_attached_in_non_strict() {
    let src = r#"
function globalFunction() { return "FuncRes"; }
console.log(globalThis.globalFunction());
"#;
    assert_eq!(run_js(src), vec!["FuncRes"]);
}

#[test]
fn test_js_globalthis_property_deletion() {
    let src = r#"
globalThis.tempVar = 999;
const deleted = delete globalThis.tempVar;
console.log(deleted + "|hasVar=" + ("tempVar" in globalThis));
"#;
    assert_eq!(run_js(src), vec!["true|hasVar=false"]);
}

#[test]
fn test_js_globalthis_property_descriptor_non_enumerable() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(globalThis, "globalThis");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["true|false|true"]);
}

#[test]
fn test_js_globalthis_in_unbound_function_call_strict_vs_non_strict() {
    let src = r#"
function nonStrictThis() { return this; }
function strictThis() { "use strict"; return this; }

console.log((nonStrictThis() === globalThis) + "|" + (strictThis() === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_globalthis_symbol_property_assignment() {
    let src = r#"
const sym = Symbol("globalSym");
globalThis[sym] = "GlobalSymData";
console.log(globalThis[sym]);
"#;
    assert_eq!(run_js(src), vec!["GlobalSymData"]);
}

#[test]
fn test_js_globalthis_preventextensions_or_freeze() {
    let src = r#"
console.log(Object.isExtensible(globalThis));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_globalthis_eval_scope_binding() {
    let src = r#"
eval("globalThis.evalVar = 'EvalData';");
console.log(evalVar);
"#;
    assert_eq!(run_js(src), vec!["EvalData"]);
}

#[test]
fn test_js_globalthis_shadowing_with_local_variable() {
    let src = r#"
const globalThisLocal = "shadowed";
console.log(globalThisLocal + "|" + (typeof globalThis.Object !== "undefined"));
"#;
    assert_eq!(run_js(src), vec!["shadowed|true"]);
}

#[test]
fn test_js_globalthis_json_stringify_handling() {
    let src = r#"
try {
    JSON.stringify(globalThis); // Cyclic reference to self throws TypeError!
} catch (e) {
    console.log("GlobalThis JSON Cycle TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["GlobalThis JSON Cycle TypeError"]);
}

#[test]
fn test_js_globalthis_is_same_value_zero_to_window_or_global() {
    let src = r#"
console.log(Object.is(globalThis, globalThis.globalThis));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_globalthis_hasown_utility() {
    let src = r#"
globalThis.ownGlobalProp = 42;
console.log(Object.hasOwn(globalThis, "ownGlobalProp"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_globalthis_reflect_ownkeys_includes_standard_builtins() {
    let src = r#"
const keys = Reflect.ownKeys(globalThis);
console.log(keys.includes("Object") && keys.includes("Array") && keys.includes("Math"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_globalthis_indirect_eval_assignment() {
    let src = r#"
(0, eval)("var indirectGlobal = 'IndirectValue';");
console.log(globalThis.indirectGlobal);
"#;
    assert_eq!(run_js(src), vec!["IndirectValue"]);
}

#[test]
fn test_js_globalthis_custom_prototype_lookup() {
    let src = r#"
console.log(Object.getPrototypeOf(globalThis) === Object.prototype || Object.getPrototypeOf(globalThis) !== null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
