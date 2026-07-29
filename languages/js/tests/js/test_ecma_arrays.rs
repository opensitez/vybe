use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Arrays — methods, iteration, modern features
// ═══════════════════════════════════════════════════════════

// ── Creation ───────────────────────────────────────────────

#[test]
fn array_literal() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
console.log(arr.length);
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn empty_array_literal_starts_empty() {
    let out = run_js(
        r#"
const arr = [];
console.log(arr.length);
arr.push(1);
arr.push(2);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["0", "1,2"]);
}

#[test]
fn function_local_array_pushes_in_loop() {
    let out = run_js(
        r#"
function main() {
    const arr = [];
    for (let i = 1; i <= 3; i++) {
        arr.push(i);
    }
    console.log(arr.join(","));
}
main();
"#,
    );
    assert_eq!(out, vec!["1,2,3"]);
}

#[test]
fn array_from() {
    let out = run_js(
        r#"
const arr = Array.from([1, 2, 3]);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["1,2,3"]);
}

#[test]
fn array_isarray() {
    let out = run_js(
        r#"
console.log(Array.isArray([1, 2]));
console.log(Array.isArray("hello"));
console.log(Array.isArray({ length: 1 }));
"#,
    );
    assert_eq!(out, vec!["true", "false", "false"]);
}

// ── Mutating methods ───────────────────────────────────────

#[test]
fn push_pop() {
    let out = run_js(
        r#"
const arr = [1, 2];
arr.push(3);
console.log(arr.length);
const last = arr.pop();
console.log(last);
console.log(arr.length);
"#,
    );
    assert_eq!(out, vec!["3", "3", "2"]);
}

#[test]
fn shift_unshift() {
    let out = run_js(
        r#"
const arr = [2, 3];
arr.unshift(1);
console.log(arr.join(","));
const first = arr.shift();
console.log(first);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["1,2,3", "1", "2,3"]);
}

#[test]
fn splice_remove() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4, 5];
arr.splice(1, 2);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["1,4,5"]);
}

#[test]
fn splice_insert() {
    let out = run_js(
        r#"
const arr = [1, 4, 5];
arr.splice(1, 0, 2, 3);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["1,2,3,4,5"]);
}

// ── Non-mutating methods ───────────────────────────────────

#[test]
fn concat() {
    let out = run_js(
        r#"
const a = [1, 2];
const b = [3, 4];
console.log(a.concat(b).join(","));
"#,
    );
    assert_eq!(out, vec!["1,2,3,4"]);
}

#[test]
fn slice() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4, 5];
console.log(arr.slice(1, 3).join(","));
console.log(arr.slice(3).join(","));
"#,
    );
    assert_eq!(out, vec!["2,3", "4,5"]);
}

#[test]
fn join() {
    let out = run_js(
        r#"
console.log([1, 2, 3].join("-"));
console.log(["a", "b", "c"].join(""));
"#,
    );
    assert_eq!(out, vec!["1-2-3", "abc"]);
}

#[test]
fn reverse() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
arr.reverse();
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["3,2,1"]);
}

#[test]
fn sort_default() {
    let out = run_js(
        r#"
const arr = [3, 1, 4, 1, 5];
arr.sort();
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["1,1,3,4,5"]);
}

#[test]
fn sort_with_comparator() {
    let out = run_js(
        r#"
const arr = [3, 1, 4, 1, 5];
arr.sort((a, b) => b - a);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["5,4,3,1,1"]);
}

#[test]
fn flat() {
    let out = run_js(
        r#"
const arr = [[1, 2], [3, 4], [5]];
console.log(arr.flat().join(","));
"#,
    );
    assert_eq!(out, vec!["1,2,3,4,5"]);
}

#[test]
fn fill() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4, 5];
arr.fill(0, 1, 4);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["1,0,0,0,5"]);
}

// ── Search methods ─────────────────────────────────────────

#[test]
fn indexof() {
    let out = run_js(
        r#"
const arr = [10, 20, 30, 20];
console.log(arr.indexOf(20));
console.log(arr.indexOf(40));
"#,
    );
    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn includes() {
    let out = run_js(
        r#"
console.log([1, 2, 3].includes(2));
console.log([1, 2, 3].includes(4));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn find() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4, 5];
const found = arr.find(x => x > 3);
console.log(found);
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn findindex() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4, 5];
const idx = arr.findIndex(x => x > 3);
console.log(idx);
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn some() {
    let out = run_js(
        r#"
console.log([1, 2, 3].some(x => x > 2));
console.log([1, 2, 3].some(x => x > 5));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn every() {
    let out = run_js(
        r#"
console.log([2, 4, 6].every(x => x % 2 === 0));
console.log([2, 3, 6].every(x => x % 2 === 0));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

// ── Iteration methods ──────────────────────────────────────

#[test]
fn map() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
const doubled = arr.map(x => x * 2);
console.log(doubled.join(","));
"#,
    );
    assert_eq!(out, vec!["2,4,6"]);
}

#[test]
fn filter() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4, 5, 6];
const evens = arr.filter(x => x % 2 === 0);
console.log(evens.join(","));
"#,
    );
    assert_eq!(out, vec!["2,4,6"]);
}

#[test]
fn reduce() {
    let out = run_js(
        r#"
const sum = [1, 2, 3, 4].reduce((acc, x) => acc + x, 0);
console.log(sum);
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn foreach() {
    let out = run_js(
        r#"
let sum = 0;
[1, 2, 3].forEach(x => { sum += x; });
console.log(sum);
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn map_filter_chain() {
    let out = run_js(
        r#"
const result = [1, 2, 3, 4, 5]
    .filter(x => x % 2 !== 0)
    .map(x => x * x);
console.log(result.join(","));
"#,
    );
    assert_eq!(out, vec!["1,9,25"]);
}

#[test]
fn reduce_to_object() {
    let out = run_js(
        r#"
const pairs = [["a", 1], ["b", 2], ["c", 3]];
const obj = pairs.reduce((acc, [k, v]) => {
    acc[k] = v;
    return acc;
}, {});
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

// ── Modern array methods ───────────────────────────────────

#[test]
fn array_at() {
    let out = run_js(
        r#"
const arr = [10, 20, 30, 40, 50];
console.log(arr.at(0));
console.log(arr.at(-1));
console.log(arr.at(-2));
"#,
    );
    assert_eq!(out, vec!["10", "50", "40"]);
}

#[test]
fn array_flat_nested() {
    let out = run_js(
        r#"
const arr = [1, [2, [3, [4]]]];
console.log(arr.flat().join(","));
"#,
    );
    assert_eq!(out, vec!["1,2,3,4"]);
}

#[test]
fn array_entries_pattern() {
    let out = run_js(
        r#"
const arr = ["a", "b", "c"];
let result = [];
for (let i = 0; i < arr.length; i++) {
    result.push(i + ":" + arr[i]);
}
console.log(result.join(","));
"#,
    );
    assert_eq!(out, vec!["0:a,1:b,2:c"]);
}

#[test]
fn slice_with_negative_indexes() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4, 5];
console.log(arr.slice(-2).join(","));
console.log(arr.slice(1, -1).join(","));
"#,
    );
    assert_eq!(out, vec!["4,5", "2,3,4"]);
}

#[test]
fn concat_does_not_mutate_original_arrays() {
    let out = run_js(
        r#"
const a = [1, 2];
const b = [3, 4];
const c = a.concat(b);
console.log(a.join(","));
console.log(b.join(","));
console.log(c.join(","));
"#,
    );
    assert_eq!(out, vec!["1,2", "3,4", "1,2,3,4"]);
}

#[test]
fn reverse_mutates_in_place() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
const same = arr.reverse();
console.log(arr === same);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["true", "3,2,1"]);
}

#[test]
fn sort_default_is_lexicographic() {
    let out = run_js(
        r#"
const arr = [10, 2, 1];
arr.sort();
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["1,10,2"]);
}

#[test]
fn fill_without_end_fills_to_array_end() {
    let out = run_js(
        r#"
const arr = [1, 2, 3, 4];
arr.fill(9, 2);
console.log(arr.join(","));
"#,
    );
    assert_eq!(out, vec!["1,2,9,9"]);
}

#[test]
fn find_returns_undefined_when_missing() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
console.log(arr.find(x => x > 10));
"#,
    );
    assert_eq!(out, vec!["undefined"]);
}

#[test]
fn findindex_returns_negative_one_when_missing() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
console.log(arr.findIndex(x => x > 10));
"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn some_short_circuits_after_match() {
    let out = run_js(
        r#"
let seen = [];
const result = [1, 2, 3, 4].some(x => {
    seen.push(x);
    return x === 3;
});
console.log(result);
console.log(seen.join(","));
"#,
    );
    assert_eq!(out, vec!["true", "1,2,3"]);
}

#[test]
fn every_short_circuits_after_failure() {
    let out = run_js(
        r#"
let seen = [];
const result = [2, 4, 5, 6].every(x => {
    seen.push(x);
    return x % 2 === 0;
});
console.log(result);
console.log(seen.join(","));
"#,
    );
    assert_eq!(out, vec!["false", "2,4,5"]);
}

#[test]
fn map_preserves_length() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
const mapped = arr.map(x => x * 10);
console.log(arr.length);
console.log(mapped.length);
"#,
    );
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn filter_can_return_empty_array() {
    let out = run_js(
        r#"
const arr = [1, 3, 5];
const even = arr.filter(x => x % 2 === 0);
console.log(even.length);
console.log(even.join(","));
"#,
    );
    assert_eq!(out, vec!["0", ""]);
}

#[test]
fn reduce_without_initial_uses_first_element() {
    let out = run_js(
        r#"
const total = [5, 6, 7].reduce((acc, x) => acc + x);
console.log(total);
"#,
    );
    assert_eq!(out, vec!["18"]);
}

#[test]
fn foreach_visits_items_in_order() {
    let out = run_js(
        r#"
let seen = [];
["a", "b", "c"].forEach(v => seen.push(v));
console.log(seen.join(","));
"#,
    );
    assert_eq!(out, vec!["a,b,c"]);
}

#[test]
fn flat_flattens_single_level_only_by_default() {
    let out = run_js(
        r#"
const arr = [1, [2, [3]]];
console.log(arr.flat().join(","));
"#,
    );
    assert_eq!(out, vec!["1,2,3"]);
}

#[test]
fn array_from_preserves_sparse_length() {
    let out = run_js(
        r#"
const sparse = [];
sparse[2] = "x";
const arr = Array.from(sparse);
console.log(arr.length);
console.log(arr[0]);
console.log(arr[2]);
"#,
    );
    assert_eq!(out, vec!["3", "undefined", "x"]);
}

#[test]
fn slice_end_beyond_length_clamps_to_length() {
    let out = run_js(
        r#"
const arr = [1, 2, 3];
console.log(arr.slice(1, 10).join(","));
"#,
    );
    assert_eq!(out, vec!["2,3"]);
}
