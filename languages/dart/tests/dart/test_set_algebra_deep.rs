//! Deep Set algebra: behavior coverage without padded numeric variants.

// Each case below covers a distinct semantic surface: overlap vs disjoint
// algebra, empty-set identity laws, order-insensitive equality, duplicate
// collapse, subset checks, chained algebra, and the cascade-sort regression
// (`setOp().toList()..sort()`). Avoid adding batch variants that only change
// literal values.

dart_cases! {
    union_merges_overlap_and_preserves_unique_elements => {
        r#"void main() {
  var a = {1, 2, 3, 4, 5};
  var b = {4, 5, 6, 7};
  var r = a.union(b).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["7", "1,2,3,4,5,6,7"]
    };

    union_with_disjoint_sets_keeps_all_members => {
        r#"void main() {
  var a = {0, 2, 4};
  var b = {1, 3, 5};
  var r = a.union(b).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["6", "0,1,2,3,4,5"]
    };

    union_with_self_is_equal_copy => {
        r#"void main() {
  var s = {for (var i = 1; i <= 8; i++) i};
  var u = s.union(s);
  print(u.length);
  print(u == s);
}"#,
        ["8", "true"]
    };

    intersection_keeps_only_shared_members => {
        r#"void main() {
  var a = {1, 2, 3, 4, 5, 6};
  var b = {4, 5, 6, 7, 8};
  var r = a.intersection(b).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["3", "4,5,6"]
    };

    intersection_of_disjoint_sets_is_empty => {
        r#"void main() {
  var a = {'a', 'b'};
  var b = {'c', 'd'};
  var r = a.intersection(b);
  print(r.isEmpty);
  print(r.length);
}"#,
        ["true", "0"]
    };

    intersection_with_self_is_equal_copy => {
        r#"void main() {
  var s = {10, 20, 30};
  var r = s.intersection(s);
  print(r.length);
  print(r == s);
}"#,
        ["3", "true"]
    };

    difference_removes_members_present_in_other => {
        r#"void main() {
  var a = {1, 2, 3, 4, 5};
  var b = {2, 4, 6};
  var r = a.difference(b).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["3", "1,3,5"]
    };

    difference_with_empty_set_returns_original_members => {
        r#"void main() {
  var a = {1, 2, 3};
  var r = a.difference(<int>{}).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["3", "1,2,3"]
    };

    difference_with_self_yields_empty_set => {
        r#"void main() {
  var s = {for (var i = 1; i <= 5; i++) i};
  var r = s.difference(s);
  print(r.isEmpty);
  print(r.length);
}"#,
        ["true", "0"]
    };

    contains_all_accepts_proper_subset => {
        r#"void main() {
  var sup = {1, 2, 3, 4, 5};
  var sub = {2, 4};
  print(sup.containsAll(sub));
}"#,
        ["true"]
    };

    contains_all_rejects_missing_member => {
        r#"void main() {
  var sup = {1, 2, 3};
  var sub = {2, 4};
  print(sup.containsAll(sub));
}"#,
        ["false"]
    };

    contains_all_empty_subset_is_true_for_empty_and_nonempty_sets => {
        r#"void main() {
  print({1, 2, 3}.containsAll(<int>{}));
  print(<int>{}.containsAll(<int>{}));
}"#,
        ["true", "true"]
    };

    set_equality_ignores_insertion_order => {
        r#"void main() {
  var a = {1, 2, 3};
  var b = {3, 2, 1};
  print(a == b);
}"#,
        ["true"]
    };

    set_equality_collapses_duplicate_literals => {
        r#"void main() {
  var a = {5};
  var b = {5, 5, 5};
  print(a == b);
  print(b.length);
}"#,
        ["true", "1"]
    };

    set_equality_rejects_same_length_different_members => {
        r#"void main() {
  var a = {1, 2, 3};
  var b = {1, 2, 4};
  print(a == b);
}"#,
        ["false"]
    };

    chained_union_then_intersection_keeps_expected_members => {
        r#"void main() {
  var a = {1, 2, 3, 4};
  var b = {3, 4, 5, 6};
  var c = {2, 3, 4, 7};
  var r = a.union(b).intersection(c).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["3", "2,3,4"]
    };

    symmetric_difference_via_union_minus_intersection => {
        r#"void main() {
  var a = {1, 2, 3, 4};
  var b = {3, 4, 5, 6};
  var r = a.union(b).difference(a.intersection(b)).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["4", "1,2,5,6"]
    };

    generated_set_comprehension_intersects_literal_range => {
        r#"void main() {
  var generated = {for (var i = 0; i < 10; i++) i * 2};
  var keep = {4, 6, 8, 9};
  var r = generated.intersection(keep).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["3", "4,6,8"]
    };

    cascade_sort_after_set_algebra_returns_sorted_list => {
        r#"void main() {
  var r = {3, 1}.union({2}).toList()..sort();
  print(r.join(','));
}"#,
        ["1,2,3"]
    };
}
