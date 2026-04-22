//! JavaScript generator functions via WASM stack switching.
//! `function*` declarations (or any function body containing `yield`)
//! are compiled with `chunk.is_generator = true`; calls return a
//! `Continuation`. `for ... of gen()` drives it via the GEN_NEXT
//! iterator protocol.

use super::helpers::run_js;

#[test]
fn generator_function_returns_continuation() {
    let out = run_js(r#"
function* gen() { yield 1; yield 2; yield 3; }
let g = gen();
console.log(g);
"#);
    // A `Continuation` object Display()s as `[continuation]`.
    assert_eq!(out, vec!["[continuation]"]);
}

#[test]
fn for_of_drives_generator_lazily() {
    let out = run_js(r#"
function* count() { yield 10; yield 20; yield 30; }
for (let v of count()) { console.log(v); }
"#);
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn generator_body_does_not_eagerly_execute() {
    let out = run_js(r#"
function* loud() {
    console.log("bad: body ran before resume");
    yield 1;
}
let g = loud();
console.log("ok");
"#);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn generator_with_arguments_yields_them() {
    let out = run_js(r#"
function* range(n) {
    let i = 0;
    while (i < n) { yield i; i = i + 1; }
}
for (let v of range(3)) { console.log(v); }
"#);
    assert_eq!(out, vec!["0", "1", "2"]);
}
