/// Array.fromAsync (ES2024) — async iterables, promises of arrays,
/// mapping function, error handling, mixed sync/async sources.

use super::helpers::run_js;

// ── from async iterable ───────────────────────────────────────────────────────

#[test]
fn from_async_with_async_generator() {
    assert_eq!(run_js(r#"
async function* asyncNums() {
    yield 1;
    yield 2;
    yield 3;
}
Array.fromAsync(asyncNums()).then(arr => console.log(arr.join(",")));
"#), vec!["1,2,3"]);
}

#[test]
fn from_async_with_array_of_promises() {
    assert_eq!(run_js(r#"
const promises = [
    Promise.resolve(10),
    Promise.resolve(20),
    Promise.resolve(30),
];
Array.fromAsync(promises).then(arr => console.log(arr.join(",")));
"#), vec!["10,20,30"]);
}

#[test]
fn from_async_with_sync_iterable() {
    assert_eq!(run_js(r#"
Array.fromAsync([1, 2, 3]).then(arr => console.log(arr.join(",")));
"#), vec!["1,2,3"]);
}

// ── mapping function ──────────────────────────────────────────────────────────

#[test]
fn from_async_with_map_function() {
    assert_eq!(run_js(r#"
async function* nums() { yield 1; yield 2; yield 3; }
Array.fromAsync(nums(), x => x * x).then(arr => console.log(arr.join(",")));
"#), vec!["1,4,9"]);
}

#[test]
fn from_async_map_function_can_return_promise() {
    assert_eq!(run_js(r#"
Array.fromAsync([1, 2, 3], x => Promise.resolve(x + 10))
    .then(arr => console.log(arr.join(",")));
"#), vec!["11,12,13"]);
}

#[test]
fn from_async_map_function_with_async_generator() {
    assert_eq!(run_js(r#"
async function* words() { yield "hello"; yield "world"; }
Array.fromAsync(words(), s => s.toUpperCase())
    .then(arr => console.log(arr.join(",")));
"#), vec!["HELLO,WORLD"]);
}

// ── returns a Promise ─────────────────────────────────────────────────────────

#[test]
fn from_async_returns_promise() {
    assert_eq!(run_js(r#"
const result = Array.fromAsync([1, 2, 3]);
console.log(result instanceof Promise);
"#), vec!["true"]);
}

#[test]
fn from_async_empty_iterable() {
    assert_eq!(run_js(r#"
Array.fromAsync([]).then(arr => {
    console.log(arr.length);
    console.log(Array.isArray(arr));
});
"#), vec!["0", "true"]);
}

// ── preserves order ───────────────────────────────────────────────────────────

#[test]
fn from_async_preserves_insertion_order() {
    assert_eq!(run_js(r#"
async function* delayed() {
    yield await Promise.resolve("a");
    yield await Promise.resolve("b");
    yield await Promise.resolve("c");
}
Array.fromAsync(delayed()).then(arr => console.log(arr.join("")));
"#), vec!["abc"]);
}

// ── error propagation ─────────────────────────────────────────────────────────

#[test]
fn from_async_rejects_on_generator_throw() {
    assert_eq!(run_js(r#"
async function* failing() {
    yield 1;
    throw new Error("boom");
}
Array.fromAsync(failing())
    .then(() => console.log("no"))
    .catch(e => console.log("caught:" + e.message));
"#), vec!["caught:boom"]);
}

// ── array-like objects ────────────────────────────────────────────────────────

#[test]
fn from_async_with_array_like() {
    assert_eq!(run_js(r#"
const arrayLike = { 0: "x", 1: "y", 2: "z", length: 3 };
Array.fromAsync(arrayLike).then(arr => console.log(arr.join(",")));
"#), vec!["x,y,z"]);
}

// ── set/map as async source ───────────────────────────────────────────────────

#[test]
fn from_async_with_set() {
    assert_eq!(run_js(r#"
const s = new Set([10, 20, 30]);
Array.fromAsync(s).then(arr => console.log(arr.join(",")));
"#), vec!["10,20,30"]);
}

#[test]
fn from_async_with_string() {
    assert_eq!(run_js(r#"
Array.fromAsync("abc").then(arr => console.log(arr.join("-")));
"#), vec!["a-b-c"]);
}

// ── async iterable protocol ───────────────────────────────────────────────────

#[test]
fn from_async_uses_symbol_async_iterator() {
    assert_eq!(run_js(r#"
const asyncIterable = {
    [Symbol.asyncIterator]() {
        let i = 0;
        const vals = [100, 200, 300];
        return {
            async next() {
                if (i < vals.length) return { value: vals[i++], done: false };
                return { done: true };
            }
        };
    }
};
Array.fromAsync(asyncIterable).then(arr => console.log(arr.join(",")));
"#), vec!["100,200,300"]);
}
