/// Control flow — labeled statements, do-while, switch fall-through, for-in
/// semantics, try/catch/finally edge cases, comma operator, void operator,
/// delete operator, nested break/continue, switch with expression types.
use super::helpers::run_js;

// ── Labeled break ─────────────────────────────────────────────────────────────

#[test]
fn labeled_break_exits_outer_loop() {
    assert_eq!(
        run_js(
            r#"
let result = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) break outer;
        result.push(i + "," + j);
    }
}
console.log(result.join("|"));
"#
        ),
        vec!["0,0"]
    );
}

#[test]
fn labeled_continue_skips_outer_iteration() {
    assert_eq!(
        run_js(
            r#"
let result = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) continue outer;
        result.push(i + "," + j);
    }
}
console.log(result.join("|"));
"#
        ),
        vec!["0,0|1,0|2,0"]
    );
}

#[test]
fn labeled_block_break() {
    assert_eq!(
        run_js(
            r#"
let x = 0;
block: {
    x = 1;
    break block;
    x = 2;
}
console.log(x);
"#
        ),
        vec!["1"]
    );
}

// ── do-while ─────────────────────────────────────────────────────────────────

#[test]
fn do_while_loops_until_condition_false() {
    assert_eq!(
        run_js(
            r#"
let n = 1;
let product = 1;
do {
    product *= n;
    n++;
} while (n <= 5);
console.log(product);
"#
        ),
        vec!["120"]
    );
}

#[test]
fn do_while_break_exits_early() {
    assert_eq!(
        run_js(
            r#"
let i = 0;
do {
    if (i === 3) break;
    i++;
} while (i < 10);
console.log(i);
"#
        ),
        vec!["3"]
    );
}

// ── switch fall-through ───────────────────────────────────────────────────────

#[test]
fn switch_falls_through_cases_without_break() {
    assert_eq!(
        run_js(
            r#"
let result = [];
switch (1) {
    case 1: result.push("one");
    case 2: result.push("two");
    case 3: result.push("three"); break;
    case 4: result.push("four");
}
console.log(result.join(","));
"#
        ),
        vec!["one,two,three"]
    );
}

#[test]
fn switch_default_at_end_runs_when_no_match() {
    assert_eq!(
        run_js(
            r#"
let x = "z";
switch (x) {
    case "a": console.log("a"); break;
    case "b": console.log("b"); break;
    default: console.log("other");
}
"#
        ),
        vec!["other"]
    );
}

#[test]
fn switch_default_in_middle_still_used() {
    assert_eq!(
        run_js(
            r#"
let result = [];
switch (99) {
    case 1: result.push("one"); break;
    default: result.push("default");
    case 2: result.push("two"); break;
}
console.log(result.join(","));
"#
        ),
        vec!["default,two"]
    );
}

#[test]
fn switch_with_string_cases() {
    assert_eq!(
        run_js(
            r#"
function grade(score) {
    switch (true) {
        case score >= 90: return "A";
        case score >= 80: return "B";
        case score >= 70: return "C";
        default: return "F";
    }
}
console.log(grade(95));
console.log(grade(82));
console.log(grade(60));
"#
        ),
        vec!["A", "B", "F"]
    );
}

#[test]
fn switch_strict_equality_no_coercion() {
    assert_eq!(
        run_js(
            r#"
let val = "1";
switch (val) {
    case 1: console.log("number"); break;
    case "1": console.log("string"); break;
}
"#
        ),
        vec!["string"]
    );
}

// ── for-in ───────────────────────────────────────────────────────────────────

#[test]
fn for_in_enumerates_own_enumerable_string_keys() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.sort().join(","));
"#
        ),
        vec!["a,b,c"]
    );
}

#[test]
fn for_in_skips_non_enumerable_properties() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperty(obj, "hidden", { value: 42, enumerable: false });
obj.visible = 1;
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
"#
        ),
        vec!["visible"]
    );
}

#[test]
fn for_in_includes_inherited_enumerable_props() {
    assert_eq!(
        run_js(
            r#"
const parent = { inherited: true };
const child = Object.create(parent);
child.own = true;
const keys = [];
for (const k in child) keys.push(k);
console.log(keys.includes("inherited"));
console.log(keys.includes("own"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn for_in_on_array_yields_indices_as_strings() {
    assert_eq!(
        run_js(
            r#"
const arr = ["x", "y", "z"];
const indices = [];
for (const i in arr) indices.push(typeof i + ":" + i);
console.log(indices.join("|"));
"#
        ),
        vec!["string:0|string:1|string:2"]
    );
}

// ── try / catch / finally edge cases ─────────────────────────────────────────

#[test]
fn finally_runs_even_when_catch_throws() {
    assert_eq!(
        run_js(
            r#"
let log = [];
try {
    try {
        throw new Error("inner");
    } catch (e) {
        log.push("caught:" + e.message);
        throw new Error("rethrown");
    } finally {
        log.push("finally");
    }
} catch (e) {
    log.push("outer:" + e.message);
}
console.log(log.join("|"));
"#
        ),
        vec!["caught:inner|finally|outer:rethrown"]
    );
}

#[test]
fn finally_return_overrides_try_return() {
    assert_eq!(
        run_js(
            r#"
function f() {
    try { return "try"; }
    finally { return "finally"; }
}
console.log(f());
"#
        ),
        vec!["finally"]
    );
}

#[test]
fn finally_always_runs_on_normal_exit() {
    assert_eq!(
        run_js(
            r#"
let log = [];
function f() {
    try { log.push("try"); return 1; }
    finally { log.push("finally"); }
}
f();
console.log(log.join(","));
"#
        ),
        vec!["try,finally"]
    );
}

#[test]
fn catch_optional_binding_omitted() {
    assert_eq!(
        run_js(
            r#"
let caught = false;
try { throw new Error("x"); } catch { caught = true; }
console.log(caught);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn nested_try_catch_inner_handled() {
    assert_eq!(
        run_js(
            r#"
let result = "none";
try {
    try { throw new RangeError("inner"); }
    catch (e) {
        if (e instanceof RangeError) result = "range";
        else throw e;
    }
} catch (e) {
    result = "outer";
}
console.log(result);
"#
        ),
        vec!["range"]
    );
}

#[test]
fn throw_non_error_value() {
    assert_eq!(
        run_js(
            r#"
let caught;
try { throw 42; } catch (e) { caught = e; }
console.log(caught);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn throw_string() {
    assert_eq!(
        run_js(
            r#"
try { throw "oops"; } catch (e) { console.log(e); }
"#
        ),
        vec!["oops"]
    );
}

// ── comma operator ────────────────────────────────────────────────────────────

#[test]
fn comma_operator_evaluates_left_to_right_returns_last() {
    assert_eq!(
        run_js(
            r#"
let x = (1, 2, 3);
console.log(x);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn comma_operator_in_for_update() {
    assert_eq!(
        run_js(
            r#"
let sum = 0;
for (let i = 0, j = 10; i < 3; i++, j--) {
    sum += j;
}
console.log(sum);
"#
        ),
        vec!["27"]
    );
}

// ── void operator ─────────────────────────────────────────────────────────────

#[test]
fn void_always_returns_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(void 0);
console.log(void "hello");
console.log(void (1 + 2));
"#
        ),
        vec!["undefined", "undefined", "undefined"]
    );
}

#[test]
fn void_used_to_call_iife_without_return() {
    assert_eq!(
        run_js(
            r#"
let side = 0;
void (function() { side = 42; })();
console.log(side);
"#
        ),
        vec!["42"]
    );
}

// ── delete operator ───────────────────────────────────────────────────────────

#[test]
fn delete_removes_object_property() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2 };
delete obj.a;
console.log(obj.a);
console.log("a" in obj);
"#
        ),
        vec!["undefined", "false"]
    );
}

#[test]
fn delete_returns_true_for_own_property() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 10 };
console.log(delete obj.x);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn delete_array_element_creates_hole() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4];
delete arr[1];
console.log(arr.length);
console.log(arr[1]);
"#
        ),
        vec!["4", "undefined"]
    );
}

#[test]
fn delete_non_existent_property_returns_true() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
console.log(delete obj.nonExistent);
"#
        ),
        vec!["true"]
    );
}

// ── for-of edge cases ─────────────────────────────────────────────────────────

#[test]
fn for_of_iterates_string_by_codepoint() {
    assert_eq!(
        run_js(
            r#"
const chars = [];
for (const ch of "abc") chars.push(ch);
console.log(chars.join("-"));
"#
        ),
        vec!["a-b-c"]
    );
}

#[test]
fn for_of_map_yields_key_value_pairs() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["a", 1], ["b", 2]]);
const pairs = [];
for (const [k, v] of m) pairs.push(k + "=" + v);
console.log(pairs.join(","));
"#
        ),
        vec!["a=1,b=2"]
    );
}

#[test]
fn for_of_set_yields_values_in_insertion_order() {
    assert_eq!(
        run_js(
            r#"
const s = new Set([3, 1, 4, 1, 5]);
const vals = [];
for (const v of s) vals.push(v);
console.log(vals.join(","));
"#
        ),
        vec!["3,1,4,5"]
    );
}

#[test]
fn for_of_break_stops_early() {
    assert_eq!(
        run_js(
            r#"
const result = [];
for (const x of [10, 20, 30, 40]) {
    if (x === 30) break;
    result.push(x);
}
console.log(result.join(","));
"#
        ),
        vec!["10,20"]
    );
}

// ── in operator ───────────────────────────────────────────────────────────────

#[test]
fn in_operator_checks_prototype_chain() {
    assert_eq!(
        run_js(
            r#"
const parent = { foo: 1 };
const child = Object.create(parent);
console.log("foo" in child);
console.log("bar" in child);
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn in_operator_with_array_indices() {
    assert_eq!(
        run_js(
            r#"
const arr = [10, 20, 30];
console.log(0 in arr);
console.log(3 in arr);
"#
        ),
        vec!["true", "false"]
    );
}

// ── short-circuit with side effects ──────────────────────────────────────────

#[test]
fn logical_and_short_circuits_right_side() {
    assert_eq!(
        run_js(
            r#"
let sideEffect = false;
false && (sideEffect = true);
console.log(sideEffect);
"#
        ),
        vec!["false"]
    );
}

#[test]
fn logical_or_short_circuits_right_side() {
    assert_eq!(
        run_js(
            r#"
let sideEffect = false;
true || (sideEffect = true);
console.log(sideEffect);
"#
        ),
        vec!["false"]
    );
}

#[test]
fn nullish_coalescing_only_for_null_and_undefined() {
    assert_eq!(
        run_js(
            r#"
console.log(0 ?? "fallback");
console.log("" ?? "fallback");
console.log(false ?? "fallback");
console.log(null ?? "fallback");
console.log(undefined ?? "fallback");
"#
        ),
        vec!["0", "", "false", "fallback", "fallback"]
    );
}

// ── while with complex conditions ─────────────────────────────────────────────

#[test]
fn while_condition_reevaluated_each_iteration() {
    assert_eq!(
        run_js(
            r#"
let arr = [1, 2, 3, 4, 5];
let sum = 0;
while (arr.length > 0) {
    sum += arr.pop();
}
console.log(sum);
"#
        ),
        vec!["15"]
    );
}

// ── throw in finally ─────────────────────────────────────────────────────────

#[test]
fn throw_in_finally_masks_try_error() {
    assert_eq!(
        run_js(
            r#"
let caught;
try {
    try { throw new Error("original"); }
    finally { throw new Error("finally"); }
} catch (e) {
    caught = e.message;
}
console.log(caught);
"#
        ),
        vec!["finally"]
    );
}

// ── continue in switch ────────────────────────────────────────────────────────

#[test]
fn continue_inside_switch_inside_loop() {
    assert_eq!(
        run_js(
            r#"
let result = [];
for (let i = 0; i < 4; i++) {
    switch (i) {
        case 2: continue;
    }
    result.push(i);
}
console.log(result.join(","));
"#
        ),
        vec!["0,1,3"]
    );
}
