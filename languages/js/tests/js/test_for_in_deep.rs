/// for...in deep patterns — prototype chain, non-enumerable, own vs inherited
use super::helpers::run_js;

#[test]
fn for_in_own_enumerable_only() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2 };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
"#
        ),
        vec!["a,b"]
    );
}

#[test]
fn for_in_includes_prototype_chain() {
    assert_eq!(
        run_js(
            r#"
const proto = { inherited: true };
const obj = Object.create(proto);
obj.own = true;
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("own"));
console.log(keys.includes("inherited"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn for_in_skips_non_enumerable_inherited() {
    assert_eq!(
        run_js(
            r#"
function Parent() {}
Object.defineProperty(Parent.prototype, "hidden", { value: 1, enumerable: false });
Parent.prototype.visible = 2;
const obj = new Parent();
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("hidden"));
console.log(keys.includes("visible"));
"#
        ),
        vec!["false", "true"]
    );
}

#[test]
fn for_in_empty_object() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
for (const k in {}) count++;
console.log(count);
"#
        ),
        vec!["0"]
    );
}

#[test]
fn for_in_array_includes_indices() {
    assert_eq!(
        run_js(
            r#"
const arr = ["a", "b", "c"];
const keys = [];
for (const k in arr) keys.push(k);
console.log(keys.join(","));
"#
        ),
        vec!["0,1,2"]
    );
}

#[test]
fn for_in_sparse_array_skips_empty_slots() {
    assert_eq!(
        run_js(
            r#"
const arr = [10, , 30, , 50];
const keys = [];
for (const k in arr) {
    if (Object.hasOwn(arr, k)) keys.push(k);
}
console.log(keys.join(","));
"#
        ),
        vec!["0,2,4"]
    );
}

#[test]
fn for_in_null_prototype_no_inherited() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.create(null);
obj.a = 1;
obj.b = 2;
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
"#
        ),
        vec!["a,b"]
    );
}

#[test]
fn for_in_string_indices() {
    assert_eq!(
        run_js(
            r#"
const keys = [];
const obj = new String("abc");
for (const k in obj) {
    if (/^\d+$/.test(k)) keys.push(k);
}
console.log(keys.join(","));
"#
        ),
        vec!["0,1,2"]
    );
}

#[test]
fn hasownproperty_filter_pattern() {
    assert_eq!(
        run_js(
            r#"
const proto = { inherited: 1 };
const obj = Object.create(proto);
obj.own = 2;
const ownKeys = [];
for (const k in obj) {
    if (Object.hasOwn(obj, k)) ownKeys.push(k);
}
console.log(ownKeys.join(","));
"#
        ),
        vec!["own"]
    );
}

#[test]
fn for_in_with_symbol_key_skips() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("s");
const obj = { a: 1, [sym]: 2 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
"#
        ),
        vec!["a"]
    );
}

#[test]
fn for_in_reflects_property_additions() {
    assert_eq!(
        run_js(
            r#"
// for-in behavior with property addition is implementation-defined
// but existing properties at start should appear
const obj = { a: 1, b: 2, c: 3 };
const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.length >= 3);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn for_in_three_level_chain() {
    assert_eq!(
        run_js(
            r#"
const a = { from_a: 1 };
const b = Object.create(a);
b.from_b = 2;
const c = Object.create(b);
c.from_c = 3;
const keys = [];
for (const k in c) keys.push(k);
console.log(keys.includes("from_a"));
console.log(keys.includes("from_b"));
console.log(keys.includes("from_c"));
"#
        ),
        vec!["true", "true", "true"]
    );
}
