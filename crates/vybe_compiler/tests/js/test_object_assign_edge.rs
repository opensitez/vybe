/// Object.assign edge cases — getters, Symbol keys, own-only, inheritance not copied

use super::helpers::run_js;

#[test]
fn assign_basic_merge() {
    assert_eq!(run_js(r#"
const result = Object.assign({}, { a: 1 }, { b: 2 }, { c: 3 });
console.log(result.a);
console.log(result.b);
console.log(result.c);
"#), vec!["1", "2", "3"]);
}

#[test]
fn assign_later_source_overwrites() {
    assert_eq!(run_js(r#"
const result = Object.assign({ x: 1 }, { x: 2 }, { x: 3 });
console.log(result.x);
"#), vec!["3"]);
}

#[test]
fn assign_copies_only_own_enumerable() {
    assert_eq!(run_js(r#"
const proto = { inherited: true };
const src = Object.create(proto);
src.own = "yes";
Object.defineProperty(src, "hidden", { value: "no", enumerable: false });
const result = Object.assign({}, src);
console.log(result.own);
console.log(result.inherited);
console.log(result.hidden);
"#), vec!["yes", "undefined", "undefined"]);
}

#[test]
fn assign_copies_symbol_keys() {
    assert_eq!(run_js(r#"
const sym = Symbol("s");
const src = { [sym]: 42, str: "ok" };
const result = Object.assign({}, src);
console.log(result[sym]);
console.log(result.str);
"#), vec!["42", "ok"]);
}

#[test]
fn assign_reads_getter_from_source() {
    assert_eq!(run_js(r#"
// assign reads the VALUE of a getter, not copies the accessor
const src = { get x() { return 42; } };
const result = Object.assign({}, src);
// result.x is a plain data property with value 42
const desc = Object.getOwnPropertyDescriptor(result, "x");
console.log(desc.value);
console.log(typeof desc.get);
"#), vec!["42", "undefined"]);
}

#[test]
fn assign_does_not_copy_getters_as_getters() {
    assert_eq!(run_js(r#"
let calls = 0;
const src = { get prop() { calls++; return calls; } };
const result = Object.assign({}, src);
const val = result.prop; // reads once already during assign
console.log(val); // the value at time of assign
console.log(calls); // only called once during assign
"#), vec!["1", "1"]);
}

#[test]
fn assign_returns_target() {
    assert_eq!(run_js(r#"
const target = { a: 1 };
const returned = Object.assign(target, { b: 2 });
console.log(returned === target);
"#), vec!["true"]);
}

#[test]
fn assign_with_null_undefined_sources_ignored() {
    assert_eq!(run_js(r#"
const result = Object.assign({ a: 1 }, null, undefined, { b: 2 });
console.log(result.a);
console.log(result.b);
"#), vec!["1", "2"]);
}

#[test]
fn assign_mutates_target() {
    assert_eq!(run_js(r#"
const target = { x: 0 };
Object.assign(target, { x: 99 });
console.log(target.x);
"#), vec!["99"]);
}

#[test]
fn assign_to_freeze_throws() {
    assert_eq!(run_js(r#"
const frozen = Object.freeze({ a: 1 });
let threw = false;
try {
    Object.assign(frozen, { b: 2 });
} catch {
    threw = true;
}
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn assign_spreads_string_chars() {
    assert_eq!(run_js(r#"
const result = Object.assign({}, "abc");
console.log(result[0]);
console.log(result[1]);
console.log(result[2]);
"#), vec!["a", "b", "c"]);
}
