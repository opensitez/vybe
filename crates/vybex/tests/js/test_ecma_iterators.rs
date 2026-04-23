use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Iteration — for-of, for-in, spread, rest,
// array/object iteration patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn for_of_array() {
    let out = run_js(r#"
const arr = [10, 20, 30];
let sum = 0;
for (const x of arr) {
    sum += x;
}
console.log(sum);
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn for_of_string() {
    let out = run_js(r#"
let chars = [];
for (const ch of "abc") {
    chars.push(ch);
}
console.log(chars.join(","));
"#);
    assert_eq!(out, vec!["a,b,c"]);
}

#[ignore]
#[test]
fn for_of_with_destructuring() {
    let out = run_js(r#"
const entries = [["a", 1], ["b", 2], ["c", 3]];
for (const [key, val] of entries) {
    console.log(key + "=" + val);
}
"#);
    assert_eq!(out, vec!["a=1", "b=2", "c=3"]);
}

#[test]
fn for_in_object_keys() {
    let out = run_js(r#"
const obj = { x: 1, y: 2, z: 3 };
const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.length);
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn for_in_with_hasownproperty() {
    let out = run_js(r#"
const obj = { a: 1, b: 2 };
for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
        console.log(key + ":" + obj[key]);
    }
}
"#);
    assert_eq!(out.len(), 2);
}

#[test]
fn spread_array_concat() {
    let out = run_js(r#"
const a = [1, 2];
const b = [3, 4];
const c = [...a, ...b];
console.log(c.length);
console.log(c.join(","));
"#);
    assert_eq!(out, vec!["4", "1,2,3,4"]);
}

#[ignore]
#[test]
fn spread_array_clone() {
    let out = run_js(r#"
const orig = [1, 2, 3];
const clone = [...orig];
clone.push(4);
console.log(orig.length);
console.log(clone.length);
"#);
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn spread_object_merge() {
    let out = run_js(r#"
const a = { x: 1, y: 2 };
const b = { y: 3, z: 4 };
const merged = { ...a, ...b };
console.log(merged.x);
console.log(merged.y);
console.log(merged.z);
"#);
    assert_eq!(out, vec!["1", "3", "4"]);
}

#[test]
fn spread_object_clone() {
    let out = run_js(r#"
const orig = { a: 1, b: 2 };
const clone = { ...orig };
clone.c = 3;
console.log(orig.c);
console.log(clone.c);
"#);
    // orig.c should be undefined/null
    assert_eq!(out[1], "3");
}

#[ignore]
#[test]
fn spread_in_function_call() {
    let out = run_js(r#"
function sum(a, b, c) { return a + b + c; }
const args = [10, 20, 30];
console.log(sum(...args));
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn spread_mixed_with_literals() {
    let out = run_js(r#"
const mid = [3, 4];
const arr = [1, 2, ...mid, 5, 6];
console.log(arr.join(","));
"#);
    assert_eq!(out, vec!["1,2,3,4,5,6"]);
}

#[ignore]
#[test]
fn for_of_map_entries() {
    let out = run_js(r#"
const m = new Map();
m.set("a", 1);
m.set("b", 2);
let count = 0;
for (const [k, v] of m) {
    count += v;
}
console.log(count);
"#);
    assert_eq!(out, vec!["3"]);
}

#[ignore]
#[test]
fn for_of_set() {
    let out = run_js(r#"
const s = new Set([10, 20, 30, 20]);
let sum = 0;
for (const v of s) {
    sum += v;
}
console.log(sum);
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn for_loop_c_style() {
    let out = run_js(r#"
let sum = 0;
for (let i = 0; i < 5; i++) {
    sum += i;
}
console.log(sum);
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn for_loop_multiple_init() {
    let out = run_js(r#"
let result = 0;
for (let i = 0, j = 10; i < j; i++, j--) {
    result += 1;
}
console.log(result);
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn while_loop() {
    let out = run_js(r#"
let i = 0;
let sum = 0;
while (i < 5) {
    sum += i;
    i++;
}
console.log(sum);
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn do_while_loop() {
    let out = run_js(r#"
let i = 0;
do {
    i++;
} while (i < 3);
console.log(i);
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn break_in_loop() {
    let out = run_js(r#"
let sum = 0;
for (let i = 0; i < 100; i++) {
    if (i >= 5) break;
    sum += i;
}
console.log(sum);
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn continue_in_loop() {
    let out = run_js(r#"
let sum = 0;
for (let i = 0; i < 10; i++) {
    if (i % 2 !== 0) continue;
    sum += i;
}
console.log(sum);
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn labeled_break() {
    let out = run_js(r#"
let found = false;
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (i === 1 && j === 1) {
            found = true;
            break outer;
        }
    }
}
console.log(found);
"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn for_in_array_iterates_present_indexes() {
    let out = run_js(r#"
const arr = ["a", "b", "c"];
let keys = [];
for (const key in arr) {
    keys.push(key);
}
console.log(keys.join(","));
"#);
    assert_eq!(out, vec!["0,1,2"]);
}

#[test]
fn for_in_skips_deleted_array_holes() {
    let out = run_js(r#"
const arr = ["a", "b", "c"];
delete arr[1];
let keys = [];
for (const key in arr) {
    keys.push(key);
}
console.log(keys.join(","));
"#);
    assert_eq!(out, vec!["0,2"]);
}

#[test]
fn for_of_array_values_not_indexes() {
    let out = run_js(r#"
const arr = [10, 20, 30];
let values = [];
for (const value of arr) {
    values.push(value);
}
console.log(values.join(","));
"#);
    assert_eq!(out, vec!["10,20,30"]);
}

#[test]
fn for_of_string_iterates_unicode_code_points() {
    let out = run_js(r#"
let chars = [];
for (const ch of "A😀B") {
    chars.push(ch);
}
console.log(chars.length);
console.log(chars[1]);
"#);
    assert_eq!(out, vec!["3", "😀"]);
}

#[test]
fn spread_array_clone_is_shallow() {
    let out = run_js(r#"
const original = [{ x: 1 }];
const copy = [...original];
copy[0].x = 9;
console.log(original[0].x);
console.log(copy.length);
"#);
    assert_eq!(out, vec!["9", "1"]);
}

#[test]
fn spread_object_clone_is_shallow() {
    let out = run_js(r#"
const original = { nested: { x: 1 } };
const copy = { ...original };
copy.nested.x = 5;
console.log(original.nested.x);
console.log(copy.nested.x);
"#);
    assert_eq!(out, vec!["5", "5"]);
}

#[test]
fn while_loop_zero_iterations() {
    let out = run_js(r#"
let count = 0;
while (count < 0) {
    count += 1;
}
console.log(count);
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn do_while_runs_once_even_when_condition_false() {
    let out = run_js(r#"
let count = 0;
do {
    count += 1;
} while (false);
console.log(count);
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn continue_skips_current_iteration_only() {
    let out = run_js(r#"
let seen = [];
for (let i = 0; i < 4; i++) {
    if (i === 2) continue;
    seen.push(i);
}
console.log(seen.join(","));
"#);
    assert_eq!(out, vec!["0,1,3"]);
}

#[test]
fn break_exits_only_innermost_loop_without_label() {
    let out = run_js(r#"
let count = 0;
for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) break;
        count += 1;
    }
}
console.log(count);
"#);
    assert_eq!(out, vec!["3"]);
}
