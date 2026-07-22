use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Object.assign()` Shallow Copy & Accessor Handling
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_assign_merging_own_enumerable_properties() {
    let src = r#"
const target = { a: 1 };
const source1 = { b: 2 };
const source2 = { c: 3 };
const res = Object.assign(target, source1, source2);
console.log(`${res.a}:${res.b}:${res.c}:${res === target}`);
"#;
    assert_eq!(run_js(src), vec!["1:2:3:true"]);
}

#[test]
fn test_js_object_assign_overrides_existing_properties() {
    let src = r#"
const target = { val: 10 };
const source = { val: 20 };
Object.assign(target, source);
console.log(target.val);
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_object_assign_ignores_non_enumerable_properties() {
    let src = r#"
const source = {};
Object.defineProperty(source, "hidden", { value: "secret", enumerable: false });
const target = Object.assign({}, source);
console.log("hidden" in target);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_object_assign_ignores_inherited_properties() {
    let src = r#"
const proto = { inherited: "protoVal" };
const source = Object.create(proto);
source.own = "ownVal";
const target = Object.assign({}, source);
console.log(`${target.own}:${"inherited" in target}`);
"#;
    assert_eq!(run_js(src), vec!["ownVal:false"]);
}

#[test]
fn test_js_object_assign_copies_symbol_properties() {
    let src = r#"
const sym = Symbol("id");
const source = { [sym]: "symbolValue" };
const target = Object.assign({}, source);
console.log(target[sym]);
"#;
    assert_eq!(run_js(src), vec!["symbolValue"]);
}

#[test]
fn test_js_object_assign_evaluates_getters_during_copy() {
    let src = r#"
const source = {
    get val() { return "GetterVal"; }
};
const target = Object.assign({}, source);
const desc = Object.getOwnPropertyDescriptor(target, "val");
console.log(target.val + "|hasGetter=" + (desc.get !== undefined));
"#;
    assert_eq!(run_js(src), vec!["GetterVal|hasGetter=false"]); // Copies value as data property, getter is NOT preserved!
}

#[test]
fn test_js_object_assign_triggers_setters_on_target() {
    let src = r#"
let targetSetterCalled = false;
const target = {
    set a(v) { targetSetterCalled = true; this._a = v * 2; }
};
Object.assign(target, { a: 10 });
console.log(targetSetterCalled + "|" + target._a);
"#;
    assert_eq!(run_js(src), vec!["true|20"]); // Triggers setter on target if property exists as accessor on target!
}

#[test]
fn test_js_object_assign_shallow_copy_nested_objects() {
    let src = r#"
const nested = { count: 1 };
const source = { inner: nested };
const target = Object.assign({}, source);
target.inner.count = 99;
console.log(source.inner.count); // Modifying target's nested object mutates source's nested object!
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_object_assign_coerces_target_to_object() {
    let src = r#"
const numObj = Object.assign(42, { a: 1 });
console.log((typeof numObj === "object") + "|" + numObj.a);
"#;
    assert_eq!(run_js(src), vec!["true|1"]);
}

#[test]
fn test_js_object_assign_null_or_undefined_target_throws_typeerror() {
    let src = r#"
try {
    Object.assign(null, { a: 1 });
} catch (e) {
    console.log("Assign Null Target TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Assign Null Target TypeError"]);
}

#[test]
fn test_js_object_assign_null_or_undefined_sources_ignored() {
    let src = r#"
const target = Object.assign({ a: 1 }, null, undefined, { b: 2 });
console.log(`${target.a}:${target.b}`);
"#;
    assert_eq!(run_js(src), vec!["1:2"]);
}

#[test]
fn test_js_object_assign_primitive_sources_coerced_only_strings_have_enumerable_own() {
    let src = r#"
const target = Object.assign({}, "abc", 123, true);
console.log(Object.values(target).join(",")); // String primitives are wrapped, indices 0,1,2 copied!
"#;
    assert_eq!(run_js(src), vec!["a,b,c"]);
}

#[test]
fn test_js_object_assign_throws_on_non_writable_target_property() {
    let src = r#"
const target = {};
Object.defineProperty(target, "fixed", { value: 10, writable: false });
try {
    Object.assign(target, { fixed: 20 });
} catch (e) {
    console.log("Assign ReadOnly Property TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Assign ReadOnly Property TypeError"]);
}

#[test]
fn test_js_object_assign_partial_application_before_exception() {
    let src = r#"
const target = {};
Object.defineProperty(target, "fixed", { value: 10, writable: false });
try {
    Object.assign(target, { a: 1, fixed: 20, b: 2 });
} catch (e) {
    console.log(`a=${target.a}|b=${target.b}`); // Property 'a' was assigned before failure on 'fixed'!
}
"#;
    assert_eq!(run_js(src), vec!["a=1|b=undefined"]);
}

#[test]
fn test_js_object_assign_array_source() {
    let src = r#"
const target = Object.assign({}, ["x", "y", "z"]);
console.log(`${target[0]}:${target[1]}:${target[2]}`);
"#;
    assert_eq!(run_js(src), vec!["x:y:z"]);
}

#[test]
fn test_js_object_assign_preserves_property_order() {
    let src = r#"
const source = { 2: "b", 1: "a", c: "c" };
const target = Object.assign({}, source);
console.log(Object.keys(target).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,c"]); // Integer indices sorted ascending, followed by string keys in insertion order!
}

#[test]
fn test_js_object_assign_property_descriptor_defaults_on_target() {
    let src = r#"
const target = Object.assign({}, { key: "val" });
const desc = Object.getOwnPropertyDescriptor(target, "key");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}`);
"#;
    assert_eq!(run_js(src), vec!["true:true:true"]); // Created properties on target are normal writable, enumerable, configurable!
}

#[test]
fn test_js_object_assign_same_target_and_source() {
    let src = r#"
const obj = { a: 1 };
const res = Object.assign(obj, obj);
console.log(res === obj + "|" + res.a);
"#;
    assert_eq!(run_js(src), vec!["true|1"]);
}

#[test]
fn test_js_object_assign_multiple_sources_evaluation_order() {
    let src = r#"
const log = [];
const s1 = { get a() { log.push("s1.a"); return 1; } };
const s2 = { get a() { log.push("s2.a"); return 2; } };
const target = Object.assign({}, s1, s2);
console.log(target.a + "|" + log.join(","));
"#;
    assert_eq!(run_js(src), vec!["2|s1.a,s2.a"]);
}

#[test]
fn test_js_object_assign_sparse_array_source() {
    let src = r#"
const sparse = [1, , 3];
const target = Object.assign({}, sparse);
console.log(`0=${target[0]}|has1=${1 in target}|2=${target[2]}`);
"#;
    assert_eq!(run_js(src), vec!["0=1|has1=false|2=3"]); // Sparse holes are non-enumerable, not copied!
}
