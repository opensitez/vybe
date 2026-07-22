use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `with` Statement & `Symbol.unscopables` Scope Exclusions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_with_statement_property_binding() {
    let src = r#"
const obj = { a: 10, b: 20 };
with (obj) {
    console.log(a + b);
}
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_symbol_unscopables_excludes_property_from_with_scope() {
    let src = r#"
const a = "OuterA";
const obj = {
    a: "InnerA",
    [Symbol.unscopables]: {
        a: true // Excludes 'a' from being bound in with(obj) scope!
    }
};
with (obj) {
    console.log(a);
}
"#;
    assert_eq!(run_js(src), vec!["OuterA"]);
}

#[test]
fn test_js_symbol_unscopables_array_prototype_defaults() {
    let src = r#"
const unscopables = Array.prototype[Symbol.unscopables];
console.log(unscopables.find + "|" + unscopables.includes + "|" + unscopables.flat);
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_with_statement_property_mutation() {
    let src = r#"
const obj = { x: 1 };
with (obj) {
    x = 100;
}
console.log(obj.x);
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_with_statement_strict_mode_prohibited_throws_syntaxerror() {
    let src = r#"
try {
    eval("'use strict'; with ({}) {}");
} catch (e) {
    console.log("With Statement Strict Mode SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["With Statement Strict Mode SyntaxError"]);
}

#[test]
fn test_js_symbol_unscopables_falsy_value_included_in_with_scope() {
    let src = r#"
const x = "OuterX";
const obj = {
    x: "InnerX",
    [Symbol.unscopables]: {
        x: false // Falsy value means 'x' IS bound in with(obj) scope!
    }
};
with (obj) {
    console.log(x);
}
"#;
    assert_eq!(run_js(src), vec!["InnerX"]);
}

#[test]
fn test_js_symbol_unscopables_null_prototype_object() {
    let src = r#"
const unscopables = Object.create(null);
unscopables.key = true;
const obj = {
    key: "InnerKey",
    [Symbol.unscopables]: unscopables
};
const key = "OuterKey";
with (obj) {
    console.log(key);
}
"#;
    assert_eq!(run_js(src), vec!["OuterKey"]);
}

#[test]
fn test_js_with_statement_prototype_chain_property_lookup() {
    let src = r#"
const proto = { inherited: "InheritedVal" };
const obj = Object.create(proto);
with (obj) {
    console.log(inherited);
}
"#;
    assert_eq!(run_js(src), vec!["InheritedVal"]);
}

#[test]
fn test_js_with_statement_expression_scope() {
    let src = r#"
function fn(o) {
    with (o) {
        return Math.max(x, y);
    }
}
console.log(fn({ x: 5, y: 15 }));
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_symbol_unscopables_well_known_symbol_identity() {
    let src = r#"
console.log(typeof Symbol.unscopables === "symbol");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_with_statement_non_existent_property_falls_through_to_outer() {
    let src = r#"
const outerVar = "OuterVal";
const obj = { a: 1 };
with (obj) {
    console.log(outerVar);
}
"#;
    assert_eq!(run_js(src), vec!["OuterVal"]);
}

#[test]
fn test_js_with_statement_assignment_to_non_existent_creates_global() {
    let src = r#"
const obj = { a: 1 };
with (obj) {
    createdInWith = "GlobalFromWith";
}
console.log(globalThis.createdInWith);
"#;
    assert_eq!(run_js(src), vec!["GlobalFromWith"]);
}

#[test]
fn test_js_symbol_unscopables_getter_on_unscopables_object() {
    let src = r#"
const val = "OuterVal";
const obj = {
    val: "InnerVal",
    [Symbol.unscopables]: {
        get val() { return true; }
    }
};
with (obj) {
    console.log(val);
}
"#;
    assert_eq!(run_js(src), vec!["OuterVal"]);
}

#[test]
fn test_js_with_statement_null_or_undefined_target_throws_typeerror() {
    let src = r#"
try {
    with (null) {}
} catch (e) {
    console.log("With Statement Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["With Statement Null TypeError"]);
}

#[test]
fn test_js_with_statement_primitive_wrapper_target() {
    let src = r#"
with ("hello") {
    console.log(length + "|" + toUpperCase());
}
"#;
    assert_eq!(run_js(src), vec!["5|HELLO"]);
}

#[test]
fn test_js_symbol_unscopables_inherited_from_prototype() {
    let src = r#"
const proto = {
    [Symbol.unscopables]: { hiddenProp: true }
};
const obj = Object.create(proto);
obj.hiddenProp = "Inner";
const hiddenProp = "Outer";
with (obj) {
    console.log(hiddenProp);
}
"#;
    assert_eq!(run_js(src), vec!["Outer"]);
}

#[test]
fn test_js_with_statement_proxy_has_trap_integration() {
    let src = r#"
const target = { a: 1 };
const proxy = new Proxy(target, {
    has(t, prop) {
        if (prop === "b") return true;
        return Reflect.has(t, prop);
    },
    get(t, prop) {
        if (prop === "b") return "TrappedB";
        return Reflect.get(t, prop);
    }
});
with (proxy) {
    console.log(a + "|" + b);
}
"#;
    assert_eq!(run_js(src), vec!["1|TrappedB"]);
}

#[test]
fn test_js_with_statement_nested_blocks() {
    let src = r#"
const o1 = { x: 1, y: 2 };
const o2 = { x: 10 };
with (o1) {
    with (o2) {
        console.log(`${x}:${y}`);
    }
}
"#;
    assert_eq!(run_js(src), vec!["10:2"]);
}

#[test]
fn test_js_symbol_unscopables_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Symbol, "unscopables");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_with_statement_function_declaration_inside_non_strict() {
    let src = r#"
const obj = {};
with (obj) {
    function fnInWith() { return "FuncInWith"; }
}
console.log(fnInWith());
"#;
    assert_eq!(run_js(src), vec!["FuncInWith"]);
}
