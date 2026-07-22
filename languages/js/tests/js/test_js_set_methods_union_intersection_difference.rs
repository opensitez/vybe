use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Set Methods (union, intersection, difference, symmetricDifference, isSubsetOf, isSupersetOf, isDisjointFrom - ES2024)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_set_union_combines_sets() {
    let src = r#"
const s1 = new Set([1, 2]);
const s2 = new Set([2, 3, 4]);
const u = s1.union(s2);
console.log([...u].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4"]);
}

#[test]
fn test_js_set_intersection_finds_common_elements() {
    let src = r#"
const s1 = new Set([1, 2, 3]);
const s2 = new Set([2, 3, 4]);
const i = s1.intersection(s2);
console.log([...i].join(","));
"#;
    assert_eq!(run_js(src), vec!["2,3"]);
}

#[test]
fn test_js_set_difference_removes_other_elements() {
    let src = r#"
const s1 = new Set([1, 2, 3, 4]);
const s2 = new Set([2, 4]);
const d = s1.difference(s2);
console.log([...d].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_set_symmetric_difference_non_overlapping_elements() {
    let src = r#"
const s1 = new Set([1, 2, 3]);
const s2 = new Set([3, 4, 5]);
const sd = s1.symmetricDifference(s2);
console.log([...sd].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,4,5"]);
}

#[test]
fn test_js_set_is_subset_of() {
    let src = r#"
const s1 = new Set([1, 2]);
const s2 = new Set([1, 2, 3]);
console.log(s1.isSubsetOf(s2) + "|" + s2.isSubsetOf(s1));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_set_is_superset_of() {
    let src = r#"
const s1 = new Set([1, 2, 3]);
const s2 = new Set([1, 2]);
console.log(s1.isSupersetOf(s2) + "|" + s2.isSupersetOf(s1));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_set_is_disjoint_from() {
    let src = r#"
const s1 = new Set([1, 2]);
const s2 = new Set([3, 4]);
const s3 = new Set([2, 3]);
console.log(s1.isDisjointFrom(s2) + "|" + s1.isDisjointFrom(s3));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_set_methods_accept_set_like_object() {
    let src = r#"
const s1 = new Set([1, 2, 3]);
const setLike = {
    size: 2,
    has(v) { return v === 2 || v === 3; },
    keys() { return [2, 3][Symbol.iterator](); }
};
const i = s1.intersection(setLike);
console.log([...i].join(","));
"#;
    assert_eq!(run_js(src), vec!["2,3"]);
}

#[test]
fn test_js_set_union_with_empty_set() {
    let src = r#"
const s1 = new Set([10, 20]);
const s2 = new Set();
console.log([...s1.union(s2)].join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20"]);
}

#[test]
fn test_js_set_intersection_with_empty_set() {
    let src = r#"
const s1 = new Set([10, 20]);
const s2 = new Set();
console.log(s1.intersection(s2).size);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_set_difference_with_identical_set() {
    let src = r#"
const s1 = new Set([1, 2]);
const s2 = new Set([1, 2]);
console.log(s1.difference(s2).size);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_set_symmetric_difference_with_identical_set() {
    let src = r#"
const s1 = new Set(["a", "b"]);
const s2 = new Set(["a", "b"]);
console.log(s1.symmetricDifference(s2).size);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_set_is_subset_of_self() {
    let src = r#"
const s = new Set([1, 2, 3]);
console.log(s.isSubsetOf(s) + "|" + s.isSupersetOf(s));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_set_is_disjoint_from_empty_set() {
    let src = r#"
const s = new Set([1, 2]);
console.log(s.isDisjointFrom(new Set()));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_set_methods_non_object_argument_throws() {
    let src = r#"
const s = new Set([1]);
try {
    s.union(12345);
} catch (e) {
    console.log("Set Method Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Set Method Non-Object TypeError"]);
}

#[test]
fn test_js_set_methods_missing_has_method_throws() {
    let src = r#"
const s = new Set([1]);
const invalidSetLike = { size: 1, keys() { return [1][Symbol.iterator](); } };
try {
    s.intersection(invalidSetLike);
} catch (e) {
    console.log("Set Method Invalid Set-Like TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Set Method Invalid Set-Like TypeError"]);
}

#[test]
fn test_js_set_methods_returns_new_set_instance() {
    let src = r#"
const s1 = new Set([1]);
const s2 = new Set([2]);
const u = s1.union(s2);
console.log((u instanceof Set) + "|" + (u !== s1) + "|" + (u !== s2));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_set_methods_object_reference_equality() {
    let src = r#"
const obj1 = { id: 1 }, obj2 = { id: 2 };
const s1 = new Set([obj1]);
const s2 = new Set([obj1, obj2]);
console.log(s1.isSubsetOf(s2));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_set_methods_map_as_set_like() {
    let src = r#"
const s = new Set([1, 2, 3]);
const map = new Map([[2, "b"], [3, "c"]]);
const i = s.intersection(map);
console.log([...i].join(","));
"#;
    assert_eq!(run_js(src), vec!["2,3"]);
}

#[test]
fn test_js_set_union_chaining() {
    let src = r#"
const s1 = new Set([1]);
const s2 = new Set([2]);
const s3 = new Set([3]);
const all = s1.union(s2).union(s3);
console.log([...all].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}
