/// structuredClone — deep copy semantics, supported types, circular references
use super::helpers::run_js;

#[test]
fn structured_clone_simple_object() {
    assert_eq!(
        run_js(
            r#"
const original = { a: 1, b: "hello", c: true };
const clone = structuredClone(original);
console.log(clone.a);
console.log(clone.b);
console.log(clone.c);
console.log(clone === original);
"#
        ),
        vec!["1", "hello", "true", "false"]
    );
}

#[test]
fn structured_clone_is_deep() {
    assert_eq!(
        run_js(
            r#"
const original = { nested: { x: 42 } };
const clone = structuredClone(original);
clone.nested.x = 99;
console.log(original.nested.x);
console.log(clone.nested.x);
"#
        ),
        vec!["42", "99"]
    );
}

#[test]
fn structured_clone_array() {
    assert_eq!(
        run_js(
            r#"
const original = [1, [2, 3], [4, [5, 6]]];
const clone = structuredClone(original);
clone[1][0] = 99;
console.log(original[1][0]);
console.log(clone[1][0]);
"#
        ),
        vec!["2", "99"]
    );
}

#[test]
fn structured_clone_map() {
    assert_eq!(
        run_js(
            r#"
const original = new Map([["key", { val: 1 }]]);
const clone = structuredClone(original);
clone.get("key").val = 99;
console.log(original.get("key").val);
console.log(clone.get("key").val);
"#
        ),
        vec!["1", "99"]
    );
}

#[test]
fn structured_clone_set() {
    assert_eq!(
        run_js(
            r#"
const original = new Set([1, 2, 3]);
const clone = structuredClone(original);
clone.add(4);
console.log(original.has(4));
console.log(clone.has(4));
console.log(clone.size);
"#
        ),
        vec!["false", "true", "4"]
    );
}

#[test]
fn structured_clone_date() {
    assert_eq!(
        run_js(
            r#"
const original = new Date(0);
const clone = structuredClone(original);
console.log(clone.getTime());
console.log(clone === original);
"#
        ),
        vec!["0", "false"]
    );
}

#[test]
fn structured_clone_regexp() {
    assert_eq!(
        run_js(
            r#"
const original = /hello/gi;
const clone = structuredClone(original);
console.log(clone.source);
console.log(clone.flags);
console.log(clone === original);
"#
        ),
        vec!["hello", "gi", "false"]
    );
}

#[test]
fn structured_clone_circular_reference() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1 };
obj.self = obj; // circular
const clone = structuredClone(obj);
console.log(clone.x);
console.log(clone.self === clone); // circular reference preserved
"#
        ),
        vec!["1", "true"]
    );
}

#[test]
fn structured_clone_preserves_reference_graph() {
    assert_eq!(
        run_js(
            r#"
const shared = { count: 0 };
const obj = { a: shared, b: shared };
const clone = structuredClone(obj);
// a and b should point to same cloned object
clone.a.count = 99;
console.log(clone.b.count); // 99 if same reference
console.log(obj.a.count);   // 0 (not mutated)
"#
        ),
        vec!["99", "0"]
    );
}

#[test]
fn structured_clone_throws_on_function() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    structuredClone({ fn: () => {} });
} catch (e) {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn structured_clone_arraybuffer() {
    assert_eq!(
        run_js(
            r#"
const buffer = new ArrayBuffer(4);
const view = new Uint8Array(buffer);
view[0] = 42;
const clonedBuffer = structuredClone(buffer);
const clonedView = new Uint8Array(clonedBuffer);
console.log(clonedView[0]);
clonedView[0] = 99;
console.log(view[0]); // original unchanged
"#
        ),
        vec!["42", "42"]
    );
}

#[test]
fn structured_clone_primitive_wrappers_not_supported() {
    assert_eq!(
        run_js(
            r#"
// Boolean, Number, String objects are cloneable
const n = new Number(42);
const clone = structuredClone(n);
console.log(+clone);
"#
        ),
        vec!["42"]
    );
}
