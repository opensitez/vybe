/// Set methods (ES2025) — union, intersection, difference, symmetricDifference,
/// isSubsetOf, isSupersetOf, isDisjointFrom.
use super::helpers::run_js;

// ── Set.prototype.union ───────────────────────────────────────────────────────

#[test]
fn set_union_combines_both_sets() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const b = new Set([3, 4, 5]);
const result = a.union(b);
console.log(result instanceof Set);
console.log([...result].sort((a,b) => a-b).join(","));
"#
        ),
        vec!["true", "1,2,3,4,5"]
    );
}

#[test]
fn set_union_with_empty_set() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const empty = new Set();
const result = a.union(empty);
console.log([...result].sort((a,b)=>a-b).join(","));
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn set_union_does_not_mutate_original() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2]);
const b = new Set([3, 4]);
const c = a.union(b);
console.log(a.size);
console.log(c.size);
"#
        ),
        vec!["2", "4"]
    );
}

// ── Set.prototype.intersection ────────────────────────────────────────────────

#[test]
fn set_intersection_common_elements() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3, 4]);
const b = new Set([3, 4, 5, 6]);
const result = a.intersection(b);
console.log([...result].sort((a,b)=>a-b).join(","));
"#
        ),
        vec!["3,4"]
    );
}

#[test]
fn set_intersection_no_common_elements() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2]);
const b = new Set([3, 4]);
console.log(a.intersection(b).size);
"#
        ),
        vec!["0"]
    );
}

#[test]
fn set_intersection_all_common() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const b = new Set([1, 2, 3]);
console.log(a.intersection(b).size);
"#
        ),
        vec!["3"]
    );
}

// ── Set.prototype.difference ──────────────────────────────────────────────────

#[test]
fn set_difference_elements_in_a_not_b() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3, 4]);
const b = new Set([3, 4, 5]);
const result = a.difference(b);
console.log([...result].sort((a,b)=>a-b).join(","));
"#
        ),
        vec!["1,2"]
    );
}

#[test]
fn set_difference_with_empty_other() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const result = a.difference(new Set());
console.log(result.size);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn set_difference_all_removed() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2]);
const result = a.difference(new Set([1, 2, 3]));
console.log(result.size);
"#
        ),
        vec!["0"]
    );
}

// ── Set.prototype.symmetricDifference ────────────────────────────────────────

#[test]
fn set_symmetric_difference_elements_in_one_only() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const b = new Set([2, 3, 4]);
const result = a.symmetricDifference(b);
console.log([...result].sort((a,b)=>a-b).join(","));
"#
        ),
        vec!["1,4"]
    );
}

#[test]
fn set_symmetric_difference_disjoint_sets() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2]);
const b = new Set([3, 4]);
const result = a.symmetricDifference(b);
console.log(result.size);
"#
        ),
        vec!["4"]
    );
}

#[test]
fn set_symmetric_difference_identical_sets() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const result = a.symmetricDifference(new Set([1, 2, 3]));
console.log(result.size);
"#
        ),
        vec!["0"]
    );
}

// ── Set.prototype.isSubsetOf ──────────────────────────────────────────────────

#[test]
fn set_is_subset_of_larger_set() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2]);
const b = new Set([1, 2, 3, 4]);
console.log(a.isSubsetOf(b));
console.log(b.isSubsetOf(a));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn set_is_subset_of_itself() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const clone = new Set([1, 2, 3]);
console.log(a.isSubsetOf(clone));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn empty_set_is_subset_of_any_set() {
    assert_eq!(
        run_js(
            r#"
const empty = new Set();
const a = new Set([1, 2, 3]);
console.log(empty.isSubsetOf(a));
const empty2 = new Set();
console.log(empty2.isSubsetOf(new Set([10, 20])));
"#
        ),
        vec!["true", "true"]
    );
}

// ── Set.prototype.isSupersetOf ────────────────────────────────────────────────

#[test]
fn set_is_superset_of_smaller_set() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3, 4]);
const b = new Set([2, 3]);
console.log(a.isSupersetOf(b));
console.log(b.isSupersetOf(a));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn set_is_superset_of_empty_set() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2]);
console.log(a.isSupersetOf(new Set()));
"#
        ),
        vec!["true"]
    );
}

// ── Set.prototype.isDisjointFrom ──────────────────────────────────────────────

#[test]
fn disjoint_sets_have_no_common_elements() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const b = new Set([4, 5, 6]);
console.log(a.isDisjointFrom(b));
console.log(b.isDisjointFrom(a));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn overlapping_sets_are_not_disjoint() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const b = new Set([3, 4, 5]);
console.log(a.isDisjointFrom(b));
"#
        ),
        vec!["false"]
    );
}

#[test]
fn empty_sets_are_disjoint() {
    assert_eq!(
        run_js(
            r#"
const a = new Set();
const b = new Set([1, 2]);
console.log(a.isDisjointFrom(b));
console.log(a.isDisjointFrom(new Set()));
"#
        ),
        vec!["true", "true"]
    );
}

// ── chaining set operations ───────────────────────────────────────────────────

#[test]
fn chain_union_then_intersection() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3]);
const b = new Set([3, 4, 5]);
const c = new Set([2, 3, 6]);
const result = a.union(b).intersection(c);
console.log([...result].sort((a,b)=>a-b).join(","));
"#
        ),
        vec!["2,3"]
    );
}
