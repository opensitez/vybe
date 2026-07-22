use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Loop Structures (`while`, `do...while`, `for`, `for...in`, `for...of`) Control Flow
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_do_while_executes_at_least_once() {
    let src = r#"
let executed = 0;
do {
    executed++;
} while (false);
console.log(executed);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_while_loop_zero_executions_when_condition_false() {
    let src = r#"
let executed = 0;
while (false) {
    executed++;
}
console.log(executed);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_for_loop_optional_expressions() {
    let src = r#"
let i = 0;
const log = [];
for (;;) {
    log.push(i);
    i++;
    if (i === 3) break;
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,2"]);
}

#[test]
fn test_js_for_in_loop_property_enumeration() {
    let src = r#"
const obj = { a: 1, b: 2 };
const props = [];
for (const prop in obj) {
    props.push(prop);
}
console.log(props.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}

#[test]
fn test_js_for_of_loop_array_elements() {
    let src = r#"
const arr = ["x", "y", "z"];
const items = [];
for (const item of arr) {
    items.push(item);
}
console.log(items.join(","));
"#;
    assert_eq!(run_js(src), vec!["x,y,z"]);
}

#[test]
fn test_js_for_of_loop_with_entries_destructuring() {
    let src = r#"
const arr = ["a", "b"];
const res = [];
for (const [idx, val] of arr.entries()) {
    res.push(`${idx}:${val}`);
}
console.log(res.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0:a|1:b"]);
}

#[test]
fn test_js_for_in_loop_skips_symbols() {
    let src = r#"
const sym = Symbol("id");
const obj = { stringProp: 1, [sym]: 2 };
const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["stringProp"]); // for...in skips Symbol keys!
}

#[test]
fn test_js_for_in_loop_includes_inherited_enumerable_properties() {
    let src = r#"
const proto = { protoProp: 10 };
const obj = Object.create(proto);
obj.ownProp = 20;
const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["ownProp,protoProp"]);
}

#[test]
fn test_js_break_statement_terminates_loop() {
    let src = r#"
const log = [];
for (let i = 0; i < 5; i++) {
    if (i === 3) break;
    log.push(i);
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,2"]);
}

#[test]
fn test_js_continue_statement_skips_iteration() {
    let src = r#"
const log = [];
for (let i = 0; i < 5; i++) {
    if (i % 2 === 0) continue;
    log.push(i);
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_for_of_loop_on_generator() {
    let src = r#"
function* gen() { yield 1; yield 2; }
const log = [];
for (const x of gen()) log.push(x);
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_for_of_loop_on_string() {
    let src = r#"
const log = [];
for (const char of "hi") log.push(char);
console.log(log.join("-"));
"#;
    assert_eq!(run_js(src), vec!["h-i"]);
}

#[test]
fn test_js_for_of_loop_non_iterable_throws_typeerror() {
    let src = r#"
try {
    for (const x of 12345);
} catch (e) {
    console.log("for...of Non-Iterable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["for...of Non-Iterable TypeError"]);
}

#[test]
fn test_js_for_in_loop_null_or_undefined_is_noop() {
    let src = r#"
let executed = false;
for (const k in null) executed = true;
for (const k in undefined) executed = true;
console.log(executed);
"#;
    assert_eq!(run_js(src), vec!["false"]); // for...in on null or undefined does 0 iterations, no error!
}

#[test]
fn test_js_for_loop_multiple_initializers_and_increments() {
    let src = r#"
const log = [];
for (let i = 0, j = 10; i < 3; i++, j -= 2) {
    log.push(`${i}:${j}`);
}
console.log(log.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0:10|1:8|2:6"]);
}

#[test]
fn test_js_do_while_condition_side_effects() {
    let src = r#"
let count = 0;
do {
    // Body empty
} while (++count < 3);
console.log(count);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_for_of_loop_iterator_return_called_on_break() {
    let src = r#"
let returned = false;
const customIterable = {
    [Symbol.iterator]() {
        return {
            next() { return { value: 1, done: false }; },
            return() { returned = true; return { done: true }; }
        };
    }
};
for (const item of customIterable) {
    break; // Breaking loop closes iterator by calling return()!
}
console.log(returned);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_for_of_loop_iterator_return_called_on_throw() {
    let src = r#"
let returned = false;
const customIterable = {
    [Symbol.iterator]() {
        return {
            next() { return { value: 1, done: false }; },
            return() { returned = true; return { done: true }; }
        };
    }
};
try {
    for (const item of customIterable) {
        throw new Error("LoopException");
    }
} catch (e) {}
console.log(returned);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_for_in_loop_deleting_property_during_enumeration() {
    let src = r#"
const obj = { a: 1, b: 2, c: 3 };
const keys = [];
for (const k in obj) {
    keys.push(k);
    if (k === "a") delete obj.b; // Deleting 'b' before visited prevents enumeration!
}
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,c"]);
}

#[test]
fn test_js_for_loop_completion_value_in_eval() {
    let src = r#"
console.log(eval("for (let i = 0; i < 2; i++) { i * 100; }"));
"#;
    assert_eq!(run_js(src), vec!["100"]);
}
