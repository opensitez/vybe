use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Array.prototype.findLast` & `findLastIndex` (ES2023)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_array_find_last_basic_predicate() {
    let src = r#"
const arr = [5, 12, 50, 130, 44];
const found = arr.findLast(x => x > 45);
console.log(found);
"#;
    assert_eq!(run_js(src), vec!["130"]);
}

#[test]
fn test_js_array_find_last_index_basic_predicate() {
    let src = r#"
const arr = [5, 12, 50, 130, 44];
const idx = arr.findLastIndex(x => x > 45);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_array_find_last_returns_undefined_when_not_found() {
    let src = r#"
const arr = [1, 2, 3];
const found = arr.findLast(x => x > 10);
console.log(found === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_array_find_last_index_returns_minus_one_when_not_found() {
    let src = r#"
const arr = [1, 2, 3];
const idx = arr.findLastIndex(x => x > 10);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["-1"]);
}

#[test]
fn test_js_array_find_last_predicate_arguments() {
    let src = r#"
const arr = ["a", "b", "c"];
const log = [];
arr.findLast((val, index, array) => {
    log.push(`${val}:${index}:${array.length}`);
    return false;
});
console.log(log.join("|"));
"#;
    assert_eq!(run_js(src), vec!["c:2:3|b:1:3|a:0:3"]); // Iterates in reverse order!
}

#[test]
fn test_js_array_find_last_this_arg_binding() {
    let src = r#"
const ctx = { threshold: 20 };
const arr = [10, 25, 30];
const found = arr.findLast(function(val) {
    return val > this.threshold;
}, ctx);
console.log(found);
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_array_find_last_sparse_array_holes_visited() {
    let src = r#"
const sparse = [1, , 3];
const visited = [];
sparse.findLast((val, idx) => {
    visited.push(`${idx}:${val}`);
    return false;
});
console.log(visited.join("|"));
"#;
    assert_eq!(run_js(src), vec!["2:3|1:undefined|0:1"]); // Sparse holes are visited as undefined!
}

#[test]
fn test_js_array_find_last_object_elements() {
    let src = r#"
const users = [
    { id: 1, active: true },
    { id: 2, active: false },
    { id: 3, active: true }
];
const lastActive = users.findLast(u => u.active);
console.log(lastActive.id);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_array_find_last_index_object_elements() {
    let src = r#"
const users = [
    { id: 1, active: true },
    { id: 2, active: false },
    { id: 3, active: true }
];
const lastActiveIdx = users.findLastIndex(u => u.active);
console.log(lastActiveIdx);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_array_find_last_array_like_object() {
    let src = r#"
const arrayLike = { 0: "first", 1: "second", length: 2 };
const found = Array.prototype.findLast.call(arrayLike, x => x.startsWith("s"));
console.log(found);
"#;
    assert_eq!(run_js(src), vec!["second"]);
}

#[test]
fn test_js_array_find_last_index_array_like_object() {
    let src = r#"
const arrayLike = { 0: "first", 1: "second", length: 2 };
const idx = Array.prototype.findLastIndex.call(arrayLike, x => x.startsWith("f"));
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_array_find_last_mutation_during_iteration() {
    let src = r#"
const arr = [1, 2, 3];
const visited = [];
arr.findLast((val) => {
    visited.push(val);
    if (val === 3) arr.pop(); // Pop element 3 during iteration
    return false;
});
console.log(visited.join(","));
"#;
    assert_eq!(run_js(src), vec!["3,2,1"]);
}

#[test]
fn test_js_array_find_last_non_callable_predicate_throws_typeerror() {
    let src = r#"
try {
    [1, 2].findLast("not_a_function");
} catch (e) {
    console.log("findLast Predicate TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["findLast Predicate TypeError"]);
}

#[test]
fn test_js_array_find_last_index_non_callable_predicate_throws_typeerror() {
    let src = r#"
try {
    [1, 2].findLastIndex(null);
} catch (e) {
    console.log("findLastIndex Predicate TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["findLastIndex Predicate TypeError"]);
}

#[test]
fn test_js_typed_array_find_last() {
    let src = r#"
const u8 = new Uint8Array([10, 20, 30, 40]);
const found = u8.findLast(x => x < 35);
console.log(found);
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_typed_array_find_last_index() {
    let src = r#"
const u8 = new Uint8Array([10, 20, 30, 40]);
const idx = u8.findLastIndex(x => x < 35);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_array_find_last_empty_array() {
    let src = r#"
const found = [].findLast(() => true);
console.log(found === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_array_find_last_index_empty_array() {
    let src = r#"
const idx = [].findLastIndex(() => true);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["-1"]);
}

#[test]
fn test_js_array_find_last_predicate_truthy_coercion() {
    let src = r#"
const arr = [0, 1, 0];
const found = arr.findLast(x => x); // Returns last truthy element
console.log(found);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_array_find_last_index_first_element_match() {
    let src = r#"
const arr = [100, 200, 300];
const idx = arr.findLastIndex(x => x === 100);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_array_find_last_index_this_arg_binding() {
    let src = r#"
const ctx = { val: 20 };
const idx = [10, 20, 30].findLastIndex(function(x) { return x === this.val; }, ctx);
console.log(idx);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

