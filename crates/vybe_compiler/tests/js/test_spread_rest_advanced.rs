use super::helpers::run_js;

// ── Spread in arrays ──────────────────────────────────────
#[test]
fn spread_concat_arrays() {
    assert_eq!(
        run_js(
            r#"
const a = [1, 2];
const b = [3, 4];
const c = [...a, ...b];
console.log(c.join(","));
"#
        ),
        vec!["1,2,3,4"]
    );
}

#[test]
fn spread_clone_array() {
    assert_eq!(
        run_js(
            r#"
const original = [1, 2, 3];
const clone = [...original];
clone.push(4);
console.log(original.length);
console.log(clone.length);
"#
        ),
        vec!["3", "4"]
    );
}

#[test]
fn spread_insert_middle() {
    assert_eq!(
        run_js(
            r#"
const start = [1, 2];
const end = [5, 6];
const mid = [3, 4];
const all = [...start, ...mid, ...end];
console.log(all.join(","));
"#
        ),
        vec!["1,2,3,4,5,6"]
    );
}

#[test]
fn spread_string_to_chars() {
    assert_eq!(
        run_js(
            r#"
const chars = [..."hello"];
console.log(chars.length);
console.log(chars[0]);
"#
        ),
        vec!["5", "h"]
    );
}

#[test]
fn spread_set_to_array() {
    assert_eq!(
        run_js(
            r#"
const set = new Set([1, 2, 3, 2, 1]);
const arr = [...set];
console.log(arr.join(","));
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn spread_map_to_array() {
    assert_eq!(
        run_js(
            r#"
const map = new Map([["a", 1], ["b", 2]]);
const arr = [...map];
console.log(arr.length);
"#
        ),
        vec!["2"]
    );
}

// ── Spread in function calls ──────────────────────────────
#[test]
fn spread_in_function_call() {
    assert_eq!(
        run_js(
            r#"
function sum(a, b, c) { return a + b + c; }
const args = [1, 2, 3];
console.log(sum(...args));
"#
        ),
        vec!["6"]
    );
}

#[test]
fn spread_in_math_max() {
    assert_eq!(
        run_js(
            r#"
const nums = [3, 1, 4, 1, 5, 9, 2, 6];
console.log(Math.max(...nums));
"#
        ),
        vec!["9"]
    );
}

#[test]
fn spread_partial_args() {
    assert_eq!(
        run_js(
            r#"
function greet(greeting, name) { return greeting + ", " + name; }
const extra = ["Alice"];
console.log(greet("Hello", ...extra));
"#
        ),
        vec!["Hello, Alice"]
    );
}

// ── Spread in objects ─────────────────────────────────────
#[test]
fn object_spread_merge() {
    assert_eq!(
        run_js(
            r#"
const defaults = { color: "red", size: 10 };
const custom = { color: "blue" };
const result = { ...defaults, ...custom };
console.log(result.color);
console.log(result.size);
"#
        ),
        vec!["blue", "10"]
    );
}

#[test]
fn object_spread_clone() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2 };
const clone = { ...obj };
clone.c = 3;
console.log(Object.keys(obj).length);
console.log(Object.keys(clone).length);
"#
        ),
        vec!["2", "3"]
    );
}

#[test]
fn object_spread_override_order() {
    assert_eq!(
        run_js(
            r#"
const base = { x: 1, y: 2, z: 3 };
const override = { x: 10 };
const merged = { ...base, ...override };
console.log(merged.x, merged.y);
"#
        ),
        vec!["10 2"]
    );
}

#[test]
fn object_spread_add_new_properties() {
    assert_eq!(
        run_js(
            r#"
const user = { name: "Alice" };
const withRole = { ...user, role: "admin" };
console.log(withRole.name);
console.log(withRole.role);
"#
        ),
        vec!["Alice", "admin"]
    );
}

#[test]
fn object_spread_nested_shallow() {
    assert_eq!(
        run_js(
            r#"
const a = { nested: { val: 1 } };
const b = { ...a };
b.nested.val = 99;
console.log(a.nested.val);
"#
        ),
        vec!["99"]
    );
}

// ── Rest parameters ───────────────────────────────────────
#[test]
fn rest_params_collects_extra() {
    assert_eq!(
        run_js(
            r#"
function log(first, ...others) {
  console.log(first);
  console.log(others.length);
}
log(1, 2, 3, 4);
"#
        ),
        vec!["1", "3"]
    );
}

#[test]
fn rest_params_can_be_empty() {
    assert_eq!(
        run_js(
            r#"
function f(a, ...b) { return b.length; }
console.log(f(1));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn rest_params_sum_variadic() {
    assert_eq!(
        run_js(
            r#"
function sum(...nums) { return nums.reduce((acc, n) => acc + n, 0); }
console.log(sum(1, 2, 3, 4, 5));
"#
        ),
        vec!["15"]
    );
}

#[test]
fn rest_params_is_real_array() {
    assert_eq!(
        run_js(
            r#"
function f(...args) { return Array.isArray(args); }
console.log(f(1, 2, 3));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn rest_params_in_arrow_function() {
    assert_eq!(
        run_js(
            r#"
const concat = (...strs) => strs.join("-");
console.log(concat("a", "b", "c"));
"#
        ),
        vec!["a-b-c"]
    );
}

// ── Object rest in destructuring ──────────────────────────
#[test]
fn object_rest_collect_remaining() {
    assert_eq!(
        run_js(
            r#"
const { a, b, ...rest } = { a: 1, b: 2, c: 3, d: 4 };
console.log(a, b);
console.log(Object.keys(rest).sort().join(","));
"#
        ),
        vec!["1 2", "c,d"]
    );
}

#[test]
fn object_rest_empty_when_all_named() {
    assert_eq!(
        run_js(
            r#"
const { x, y, ...rest } = { x: 1, y: 2 };
console.log(Object.keys(rest).length);
"#
        ),
        vec!["0"]
    );
}

// ── Spread with new ───────────────────────────────────────
#[test]
fn spread_in_new_constructor() {
    assert_eq!(
        run_js(
            r#"
const args = [2024, 0, 1];
const d = new Date(...args);
console.log(d.getFullYear());
"#
        ),
        vec!["2024"]
    );
}

// ── Spread with generators ────────────────────────────────
#[test]
fn spread_generator_into_array() {
    assert_eq!(
        run_js(
            r#"
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const arr = [...range(5)];
console.log(arr.join(","));
"#
        ),
        vec!["0,1,2,3,4"]
    );
}

#[test]
fn spread_nested_arrays_flattens_one_level() {
    assert_eq!(
        run_js(
            r#"
const nested = [[1, 2], [3, 4]];
const flat = [].concat(...nested);
console.log(flat.join(","));
"#
        ),
        vec!["1,2,3,4"]
    );
}

#[test]
fn rest_and_spread_roundtrip() {
    assert_eq!(
        run_js(
            r#"
function pass(...args) { return args; }
const src = [1, 2, 3];
const result = pass(...src);
console.log(result.join(","));
"#
        ),
        vec!["1,2,3"]
    );
}
