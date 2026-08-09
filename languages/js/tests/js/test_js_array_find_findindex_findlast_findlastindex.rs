use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Array Searching Methods (find, findIndex, findLast, findLastIndex - ES2023)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_array_find_match_found() {
    let src = r#"
const numbers = [5, 12, 8, 130, 44];
const found = numbers.find(element => element > 10);
console.log(found);
"#;
    assert_eq!(run_js(src), vec!["12"]);
}

#[test]
fn test_js_array_find_match_not_found_returns_undefined() {
    let src = r#"
const numbers = [5, 2, 8];
const found = numbers.find(element => element > 10);
console.log(found === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_array_findindex_match_found() {
    let src = r#"
const numbers = [5, 12, 8, 130, 44];
const idx = numbers.findIndex(element => element > 10);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_array_findindex_match_not_found_returns_minus_one() {
    let src = r#"
const numbers = [5, 2, 8];
const idx = numbers.findIndex(element => element > 10);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["-1"]);
}

#[test]
fn test_js_array_findlast_match_found_es2023() {
    let src = r#"
const numbers = [5, 12, 50, 130, 44];
const lastFound = numbers.findLast(element => element > 10);
console.log(lastFound);
"#;
    assert_eq!(run_js(src), vec!["44"]);
}

#[test]
fn test_js_array_findlastindex_match_found_es2023() {
    let src = r#"
const numbers = [5, 12, 50, 130, 44];
const lastIdx = numbers.findLastIndex(element => element > 10);
console.log(lastIdx);
"#;
    assert_eq!(run_js(src), vec!["4"]);
}

#[test]
fn test_js_array_findlast_not_found_returns_undefined() {
    let src = r#"
const numbers = [1, 2, 3];
console.log(numbers.findLast(x => x > 10) === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_array_findlastindex_not_found_returns_minus_one() {
    let src = r#"
const numbers = [1, 2, 3];
console.log(numbers.findLastIndex(x => x > 10));
"#;
    assert_eq!(run_js(src), vec!["-1"]);
}

#[test]
fn test_js_array_find_predicate_arguments() {
    let src = r#"
const items = ["a"];
items.find((val, idx, arr) => {
    console.log(`${val}:${idx}:${arr === items}`);
});
"#;
    assert_eq!(run_js(src), vec!["a:0:true"]);
}

#[test]
fn test_js_array_find_this_argument_binding() {
    let src = r#"
const ctx = { threshold: 20 };
const nums = [10, 25, 30];
const res = nums.find(function(x) { return x > this.threshold; }, ctx);
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["25"]);
}

#[test]
fn test_js_array_find_sparse_array_holes_visited_as_undefined() {
    let src = r#"
const sparse = [1, , 3];
const visited = [];
sparse.find(val => {
    visited.push(String(val));
    return false;
});
console.log(visited.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,undefined,3"]);
}

#[test]
fn test_js_array_findlast_sparse_array_holes_visited_as_undefined() {
    let src = r#"
const sparse = [1, , 3];
const visited = [];
sparse.findLast(val => {
    visited.push(String(val));
    return false;
});
console.log(visited.join(","));
"#;
    assert_eq!(run_js(src), vec!["3,undefined,1"]);
}

#[test]
fn test_js_array_find_object_by_property() {
    let src = r#"
const inventory = [
    { name: "apples", quantity: 2 },
    { name: "bananas", quantity: 0 },
    { name: "cherries", quantity: 5 }
];
const result = inventory.find(item => item.name === "cherries");
console.log(result.quantity);
"#;
    assert_eq!(run_js(src), vec!["5"]);
}

#[test]
fn test_js_array_findindex_object_by_property() {
    let src = r#"
const inventory = [
    { name: "apples", quantity: 2 },
    { name: "bananas", quantity: 0 }
];
const idx = inventory.findIndex(item => item.name === "bananas");
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_array_find_stops_at_first_truthy_match() {
    let src = r#"
let calls = 0;
const nums = [1, 2, 3, 4, 5];
nums.find(x => {
    calls++;
    return x === 2;
});
console.log(calls);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_array_findlast_stops_at_first_match_from_end() {
    let src = r#"
let calls = 0;
const nums = [1, 2, 3, 4, 5];
nums.findLast(x => {
    calls++;
    return x === 4;
});
console.log(calls);
"#;
    assert_eq!(run_js(src), vec!["2"]); // Visited index 4, then index 3 (match!) -> 2 calls
}

#[test]
fn test_js_array_find_mutation_during_search() {
    let src = r#"
const nums = [0, 1, 2];
const res = nums.find((x, idx, a) => {
    if (idx === 0) a[1] = 99; // Mutates index 1 before visited
    return x === 99;
});
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_array_find_non_callable_predicate_throws() {
    let src = r#"
try {
    [1, 2].find(123);
} catch (e) {
    console.log("Find Non-Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Find Non-Callable TypeError"]);
}

#[test]
fn test_js_array_findindex_nan_matching() {
    let src = r#"
const arr = [1, NaN, 2];
const idx = arr.findIndex(Number.isNaN);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_array_findlastindex_nan_matching() {
    let src = r#"
const arr = [NaN, 1, NaN, 2];
const idx = arr.findLastIndex(Number.isNaN);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_array_findlast_this_argument_binding() {
    let src = r#"
const ctx = { cap: 30 };
const nums = [10, 20, 40, 25];
console.log(nums.findLast(function(x) { return x < this.cap; }, ctx));
"#;
    assert_eq!(run_js(src), vec!["25"]);
}
