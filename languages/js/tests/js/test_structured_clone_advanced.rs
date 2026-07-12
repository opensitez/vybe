use super::helpers::run_js;

// ── structuredClone basics ────────────────────────────────
#[test]
fn structuredclone_primitive_number() {
    assert_eq!(
        run_js(
            r#"
const n = structuredClone(42);
console.log(n);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn structuredclone_string() {
    assert_eq!(
        run_js(
            r#"
const s = structuredClone("hello");
console.log(s);
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn structuredclone_plain_object() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: "two", c: true };
const clone = structuredClone(obj);
console.log(clone.a, clone.b, clone.c);
"#
        ),
        vec!["1 two true"]
    );
}

#[test]
fn structuredclone_deep_copy() {
    assert_eq!(
        run_js(
            r#"
const orig = { nested: { x: 1 } };
const clone = structuredClone(orig);
clone.nested.x = 99;
console.log(orig.nested.x);
"#
        ),
        vec!["1"]
    );
}

#[test]
fn structuredclone_array() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, [2, 3], [4, [5]]];
const clone = structuredClone(arr);
clone[1][0] = 99;
console.log(arr[1][0]);
console.log(clone[1][0]);
"#
        ),
        vec!["2", "99"]
    );
}

#[test]
fn structuredclone_map() {
    assert_eq!(
        run_js(
            r#"
const orig = new Map([["a", 1], ["b", 2]]);
const clone = structuredClone(orig);
clone.set("c", 3);
console.log(orig.size);
console.log(clone.size);
"#
        ),
        vec!["2", "3"]
    );
}

#[test]
fn structuredclone_set() {
    assert_eq!(
        run_js(
            r#"
const orig = new Set([1, 2, 3]);
const clone = structuredClone(orig);
clone.add(4);
console.log(orig.size);
console.log(clone.size);
"#
        ),
        vec!["3", "4"]
    );
}

#[test]
fn structuredclone_date() {
    assert_eq!(
        run_js(
            r#"
const d = new Date(2024, 0, 15);
const clone = structuredClone(d);
console.log(clone instanceof Date);
console.log(clone.getFullYear());
"#
        ),
        vec!["true", "2024"]
    );
}

#[test]
fn structuredclone_regexp() {
    assert_eq!(
        run_js(
            r#"
const re = /hello/gi;
const clone = structuredClone(re);
console.log(clone instanceof RegExp);
console.log(clone.flags.includes("g"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn structuredclone_nested_maps_and_sets() {
    assert_eq!(
        run_js(
            r#"
const orig = { m: new Map([["k", new Set([1, 2])]]) };
const clone = structuredClone(orig);
clone.m.get("k").add(3);
console.log(orig.m.get("k").size);
console.log(clone.m.get("k").size);
"#
        ),
        vec!["2", "3"]
    );
}

#[test]
fn structuredclone_boolean() {
    assert_eq!(
        run_js(
            r#"
console.log(structuredClone(true));
console.log(structuredClone(false));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn structuredclone_null() {
    assert_eq!(
        run_js(
            r#"
console.log(structuredClone(null) === null);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn structuredclone_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(structuredClone(undefined) === undefined);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn structuredclone_circular_reference_throws() {
    // Per the HTML structured-clone spec, a circular reference is PRESERVED,
    // not rejected: the clone's back-edge points at the clone itself. (This
    // corrects the old assertion that it throws — browsers do not throw here.)
    assert_eq!(
        run_js(
            r#"
const obj = {};
obj.self = obj;
const clone = structuredClone(obj);
console.log(clone !== obj);
console.log(clone.self === clone);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn structuredclone_typed_array() {
    assert_eq!(
        run_js(
            r#"
const orig = new Uint8Array([1, 2, 3]);
const clone = structuredClone(orig);
clone[0] = 99;
console.log(orig[0]);
console.log(clone[0]);
"#
        ),
        vec!["1", "99"]
    );
}

#[test]
fn structuredclone_array_with_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, , 3];
const clone = structuredClone(arr);
console.log(clone.length);
console.log(clone[0], clone[2]);
"#
        ),
        vec!["3", "1 3"]
    );
}

#[test]
fn structuredclone_complex_object_graph() {
    assert_eq!(
        run_js(
            r#"
const orig = {
  users: [{ name: "Alice", scores: [1, 2, 3] }, { name: "Bob", scores: [4, 5] }],
  meta: { total: 2 }
};
const clone = structuredClone(orig);
clone.users[0].scores.push(99);
clone.meta.total = 99;
console.log(orig.users[0].scores.length);
console.log(orig.meta.total);
"#
        ),
        vec!["3", "2"]
    );
}

#[test]
fn structuredclone_preserves_array_type() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
const clone = structuredClone(arr);
console.log(Array.isArray(clone));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn structuredclone_number_zero_negative() {
    assert_eq!(
        run_js(
            r#"
console.log(structuredClone(-0) === 0);
console.log(1 / structuredClone(-0));
"#
        ),
        vec!["true", "-Infinity"]
    );
}

#[test]
fn structuredclone_infinity_nan() {
    assert_eq!(
        run_js(
            r#"
console.log(structuredClone(Infinity));
console.log(isNaN(structuredClone(NaN)));
"#
        ),
        vec!["Infinity", "true"]
    );
}
