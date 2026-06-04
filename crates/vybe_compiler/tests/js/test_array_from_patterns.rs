/// Array.from patterns — from iterable, from array-like, with mapFn, Set, Map, string, generator
use super::helpers::run_js;

#[test]
fn array_from_string() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from("hello");
console.log(arr.join(","));
"#
        ),
        vec!["h,e,l,l,o"]
    );
}

#[test]
fn array_from_set() {
    assert_eq!(
        run_js(
            r#"
const s = new Set([1, 2, 3, 2, 1]);
const arr = Array.from(s);
console.log(arr.join(","));
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn array_from_map_entries() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["a", 1], ["b", 2]]);
const arr = Array.from(m);
console.log(arr.map(([k, v]) => k + "=" + v).join(","));
"#
        ),
        vec!["a=1,b=2"]
    );
}

#[test]
fn array_from_array_like_object() {
    assert_eq!(
        run_js(
            r#"
const arrayLike = { 0: "x", 1: "y", 2: "z", length: 3 };
const arr = Array.from(arrayLike);
console.log(arr.join(","));
"#
        ),
        vec!["x,y,z"]
    );
}

#[test]
fn array_from_with_map_fn() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from([1, 2, 3], x => x * 2);
console.log(arr.join(","));
"#
        ),
        vec!["2,4,6"]
    );
}

#[test]
fn array_from_map_fn_with_index() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from("abc", (c, i) => i + ":" + c);
console.log(arr.join(","));
"#
        ),
        vec!["0:a,1:b,2:c"]
    );
}

#[test]
fn array_from_generator() {
    assert_eq!(
        run_js(
            r#"
function* range(n) {
    for (let i = 0; i < n; i++) yield i;
}
const arr = Array.from(range(5));
console.log(arr.join(","));
"#
        ),
        vec!["0,1,2,3,4"]
    );
}

#[test]
fn array_from_length_property_fills_undefined() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from({ length: 3 });
console.log(arr.length);
console.log(arr.every(x => x === undefined));
"#
        ),
        vec!["3", "true"]
    );
}

#[test]
fn array_from_length_with_map() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from({ length: 5 }, (_, i) => i * i);
console.log(arr.join(","));
"#
        ),
        vec!["0,1,4,9,16"]
    );
}

#[test]
fn array_from_empty_iterable() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from([]);
console.log(arr.length);
"#
        ),
        vec!["0"]
    );
}

#[test]
fn array_from_custom_iterator() {
    assert_eq!(
        run_js(
            r#"
const iterable = {
    [Symbol.iterator]() {
        let n = 0;
        return { next() { return n < 3 ? { value: n++, done: false } : { done: true }; } };
    }
};
const arr = Array.from(iterable);
console.log(arr.join(","));
"#
        ),
        vec!["0,1,2"]
    );
}

#[test]
fn array_of_vs_array_from() {
    assert_eq!(
        run_js(
            r#"
// Array.of: treats args as elements
const a = Array.of(7);
console.log(a.length);
console.log(a[0]);
// Array(7): creates hole array of length 7
const b = Array(7);
console.log(b.length);
console.log(b[0]);
"#
        ),
        vec!["1", "7", "7", "undefined"]
    );
}
