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
fn test_js_for_in_loop_includes_custom_own_property() {
    let src = r#"
const obj = { 0: 10, 1: 20 };
obj.extra = 30;
const keys = [];
for (const k in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, k)) {
        keys.push(k);
    }
}
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,extra"]);
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
fn test_js_labeled_break_exits_outer_loop() {
    let src = r#"
const seen = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (i === 1) {
            break outer;
        }
        seen.push(i + ":" + j);
    }
}
console.log(seen.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0:0|0:1|0:2"]);
}

#[test]
fn test_js_labeled_continue_skips_to_outer_loop() {
    let src = r#"
const seen = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) {
            continue outer;
        }
        seen.push(i + ":" + j);
    }
}
console.log(seen.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0:0|1:0|2:0"]);
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
fn test_js_for_of_break_invokes_iterator_return() {
    let src = r#"
let nextCount = 0;
let returnCount = 0;

const iterable = {
    [Symbol.iterator]() {
        return {
            next() {
                nextCount++;
                return nextCount <= 3
                    ? { value: nextCount, done: false }
                    : { done: true };
            },
            return() {
                returnCount++;
                return { done: true };
            }
        };
    }
};

const values = [];
for (const value of iterable) {
    values.push(value);
    if (value === 1) {
        break;
    }
}

console.log(values.join(","));
console.log(`${nextCount}:${returnCount}`);
"#;

    assert_eq!(run_js(src), vec!["1", "1:1"]);
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
fn test_js_for_of_loop_sparse_array_includes_undefined_hole() {
    let src = r#"
const seen = [];
for (const value of [1, , 3]) {
    seen.push(value === undefined ? "u" : String(value));
}
console.log(seen.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,u,3"]);
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
fn test_js_do_while_continue_still_checks_condition() {
    let src = r#"
let log = [];
let i = 0;
do {
    i++;
    log.push("iter");
    if (i === 1) {
        log.push("continue");
        continue;
    }
    log.push("after-" + i);
} while (i < 3);
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["iter,continue,iter,after-2,iter,after-3"]);
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
fn test_js_for_of_loop_iterator_return_not_called_on_continue() {
    let src = r#"
let returned = false;
const customIterable = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() {
                return i < 4 ? { value: ++i, done: false } : { done: true };
            },
            return() {
                returned = true;
                return { done: true };
            }
        };
    }
};
const seen = [];
for (const n of customIterable) {
    if (n % 2 === 0) continue;
    seen.push(n);
}
console.log(seen.join(","));
console.log(returned);
"#;
    assert_eq!(run_js(src), vec!["1,3", "false"]);
}

#[test]
fn test_js_for_of_loop_return_calls_iterator_return() {
    let src = r#"
let returned = false;
const customIterable = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() {
                return i < 4 ? { value: ++i, done: false } : { done: true };
            },
            return() {
                returned = true;
                return { done: true };
            }
        };
    }
};
(function consume() {
    for (const n of customIterable) {
        if (n === 2) return n;
    }
}());
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
fn test_js_for_in_loop_string_primitive_yields_indices() {
    let src = r#"
const keys = [];
for (const k in "ab") keys.push(k);
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1"]);
}

#[test]
fn test_js_for_loop_completion_value_in_eval() {
    let src = r#"
console.log(eval("for (let i = 0; i < 2; i++) { i * 100; }"));
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_while_loop_finally_runs_before_break() {
    let src = r#"
let log = [];
let i = 0;
while (i < 3) {
    try {
        log.push("try" + i);
        if (i === 1) {
            break;
        }
        i++;
    } finally {
        log.push("finally" + i);
    }
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["try0,finally1,try1,finally1"]);
}

#[test]
fn test_js_for_of_loop_continue_skips_iteration() {
    let src = r#"
const out = [];
for (const n of [1, 2, 3, 4]) {
    if (n % 2 === 0) continue;
    out.push(n);
}
console.log(out.join("|"));
"#;
    assert_eq!(run_js(src), vec!["1|3"]);
}

#[test]
fn test_js_nested_loops_with_labeled_break() {
    let src = r#"
const seen = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 0) continue;
        if (j === 2) break outer;
        seen.push(i + ":" + j);
    }
}
console.log(seen.join(","));
"#;
    assert_eq!(run_js(src), vec!["0:1"]);
}

#[test]
fn test_js_while_loop_continue_skips_body_then_still_checks_condition() {
    let src = r#"
const values = [];
let i = 0;
while (i < 5) {
    i++;
    if (i === 2) {
        continue;
    }
    if (i === 4) {
        break;
    }
    values.push(i);
}
console.log(values.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_while_loop_update_skips_work_when_continue_hits() {
    let src = r#"
let i = 0;
const values = [];
while (i < 5) {
    i += 1;
    if (i === 2) {
        continue;
    }
    values.push(i);
}
console.log(values.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3,4,5"]);
}

#[test]
fn test_js_for_loop_update_still_runs_on_continue() {
    let src = r#"
const values = [];
for (let i = 0; i < 4; i++) {
    if (i === 2) {
        continue;
    }
    values.push(i);
}
console.log(values.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,3"]);
}

#[test]
fn test_js_do_while_break_before_next_iteration() {
    let src = r#"
let i = 0;
let sum = 0;
do {
    i++;
    if (i === 2) {
        continue;
    }
    if (i === 4) {
        break;
    }
    sum += i;
} while (i < 5);
    console.log(`${sum}:${i}`);
"#;
    assert_eq!(run_js(src), vec!["4:4"]);
}

#[test]
fn test_js_for_loop_with_let_captures_per_iteration() {
    let src = r#"
const values = [];
const fns = [];
for (let i = 0; i < 3; i++) {
    fns.push(() => i);
    values.push(i);
}
console.log(values.join(","));
console.log(fns.map((fn) => fn()).join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,2", "0,1,2"]);
}

#[test]
fn test_js_do_while_finally_runs_on_continue() {
    let src = r#"
let log = [];
let i = 0;
do {
    try {
        log.push("body" + i);
        if (i === 1) {
            i++;
            continue;
        }
        i++;
    } finally {
        log.push("finally" + i);
    }
} while (i < 4);
console.log(log.join("|"));
"#;
    assert_eq!(run_js(src), vec!["body0|finally1|body1|finally2|body2|finally3|body3|finally4"]);
}

#[test]
fn test_js_do_while_with_continue_in_finally_scope() {
    let src = r#"
let log = [];
let i = 0;
do {
    try {
        log.push("body" + i);
        if (i === 0) {
            i++;
            continue;
        }
        if (i === 1) {
            break;
        }
        i++;
    } finally {
        log.push("finally" + i);
    }
} while (i < 5);
console.log(log.join("|"));
"#;
    assert_eq!(run_js(src), vec!["body0|finally1|body1|finally1"]);
}

#[test]
fn test_js_labeled_do_while_with_continue_and_break() {
    let src = r#"
let out = [];
let i = 0;
outer: do {
    i++;
    if (i === 2) {
        out.push("continue-" + i);
        continue outer;
    }
    if (i === 4) {
        out.push("break-" + i);
        break;
    }
    out.push("body-" + i);
} while (i < 5);
console.log(out.join("|"));
    "#;
    assert_eq!(run_js(src), vec!["body-1|continue-2|body-3|break-4"]);
}

#[test]
fn test_js_while_loop_condition_side_effects_and_continue() {
    let src = r#"
let i = 0;
let checks = 0;
while ((checks++, i < 3)) {
    if (i === 1) {
        i++;
        continue;
    }
    i++;
}
console.log(i + "|" + checks);
"#;
assert_eq!(run_js(src), vec!["3|4"]);
}

#[test]
fn test_js_while_loop_finally_runs_on_continue() {
    let src = r#"
const log = [];
let i = 0;
while (i < 3) {
    try {
        log.push("try-" + i);
        if (i === 1) {
            i++;
            continue;
        }
        i++;
    } finally {
        log.push("finally-" + i);
    }
}
console.log(log.join(","));
"#;
    assert_eq!(
        run_js(src),
        vec!["try-0,finally-1,try-1,finally-2,try-2,finally-3"]
    );
}

#[test]
fn test_js_for_of_destructuring_defaults_in_loop_header() {
    let src = r#"
const out = [];
for (const [a = 10, b = 20] of [[0, undefined], [null, 5]]) {
    out.push(`${a}:${b}`);
}
console.log(out.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0:20|null:5"]);
}

#[test]
fn test_js_for_of_map_destructuring() {
    let src = r#"
const m = new Map([["k1", "v1"]]);
const out = [];
for (const [k, v] of m) out.push(`${k}=${v}`);
console.log(out.join(","));
"#;
    assert_eq!(run_js(src), vec!["k1=v1"]);
}

