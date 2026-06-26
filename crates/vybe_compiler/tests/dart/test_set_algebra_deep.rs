//! Deep Set algebra: large unions/intersections/differences, disjoint sets, containsAll subset checks, and equality.

dart_cases! {
    large_union_merges_distinct_elements_batch_0 => {
        r#"void main() {
  var a = {1,2,3,4,5,6,7,8,9,10};
  var b = {6,7,8,9,10,11,12,13,14,15};
  var c = a.union(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["15", "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15"]
    };

    large_union_merges_distinct_elements_batch_1 => {
        r#"void main() {
  var a = {1,2,3,4,5};
  var b = {1,2,3,4,5};
  var c = a.union(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["5", "1,2,3,4,5"]
    };

    large_union_merges_distinct_elements_batch_2 => {
        r#"void main() {
  var a = {1,2,3,4,5,6,7};
  var b = {8,9,10,11,12,13,14};
  var c = a.union(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["14", "1,2,3,4,5,6,7,8,9,10,11,12,13,14"]
    };

    large_union_merges_distinct_elements_batch_3 => {
        r#"void main() {
  var a = {10,20,30,40,50};
  var b = {30,40,50,60,70};
  var c = a.union(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["7", "10,20,30,40,50,60,70"]
    };

    large_union_merges_distinct_elements_batch_4 => {
        r#"void main() {
  var a = {0,2,4,6,8,10,12,14,16,18};
  var b = {1,3,5,7,9,11,13,15,17,19};
  var c = a.union(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["20", "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19"]
    };

    large_intersection_keeps_shared_elements_batch_0 => {
        r#"void main() {
  var a = {1,2,3,4,5,6,7,8,9,10};
  var b = {5,6,7,8,9,10,11,12,13,14};
  var c = a.intersection(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["6", "5,6,7,8,9,10"]
    };

    large_intersection_keeps_shared_elements_batch_1 => {
        r#"void main() {
  var a = {1,2,3,4,5};
  var b = {10,11,12,13,14};
  var c = a.intersection(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["0", ""]
    };

    large_intersection_keeps_shared_elements_batch_2 => {
        r#"void main() {
  var a = {1,3,5,7,9,11};
  var b = {2,3,5,7,11,13};
  var c = a.intersection(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["4", "3,5,7,11"]
    };

    large_intersection_keeps_shared_elements_batch_3 => {
        r#"void main() {
  var a = {100,101,102,103,104,105,106,107,108,109};
  var b = {105,106,107,108,109,110,111,112,113,114};
  var c = a.intersection(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["5", "105,106,107,108,109"]
    };

    large_intersection_keeps_shared_elements_batch_4 => {
        r#"void main() {
  var a = {2,4,6,8};
  var b = {1,2,3,4};
  var c = a.intersection(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["2", "2,4"]
    };

    large_difference_removes_other_elements_batch_0 => {
        r#"void main() {
  var a = {1,2,3,4,5,6,7,8,9,10};
  var b = {5,6,7,8,9,10,11,12,13,14};
  var c = a.difference(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["4", "1,2,3,4"]
    };

    large_difference_removes_other_elements_batch_1 => {
        r#"void main() {
  var a = {1,2,3,4,5,6,7,8,9,10};
  var b = {1,2,3,4,5};
  var c = a.difference(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["5", "6,7,8,9,10"]
    };

    large_difference_removes_other_elements_batch_2 => {
        r#"void main() {
  var a = {10,20,30,40,50};
  var b = {20,40,60};
  var c = a.difference(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["3", "10,30,50"]
    };

    large_difference_removes_other_elements_batch_3 => {
        r#"void main() {
  var a = {1,2,3,4,5,6,7};
  var b = <int>{};
  var c = a.difference(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["7", "1,2,3,4,5,6,7"]
    };

    large_difference_removes_other_elements_batch_4 => {
        r#"void main() {
  var a = {1,2,3};
  var b = {1,2,3,4,5};
  var c = a.difference(b).toList()..sort();
  print(c.length);
  print(c.join(','));
}"#,
        ["0", ""]
    };

    disjoint_sets_intersection_is_empty_batch_0 => {
        r#"void main() {
  var a = {1,2,3,4,5};
  var b = {10,11,12,13,14};
  var c = a.intersection(b);
  print(c.isEmpty);
  print(c.length);
}"#,
        ["true", "0"]
    };

    disjoint_sets_intersection_is_empty_batch_1 => {
        r#"void main() {
  var a = {100,101,102,103,104};
  var b = {200,201,202,203,204};
  var c = a.intersection(b);
  print(c.isEmpty);
  print(c.length);
}"#,
        ["true", "0"]
    };

    disjoint_sets_intersection_is_empty_batch_2 => {
        r#"void main() {
  var a = {1,3,5};
  var b = {2,4,6};
  var c = a.intersection(b);
  print(c.isEmpty);
  print(c.length);
}"#,
        ["true", "0"]
    };

    disjoint_sets_intersection_is_empty_batch_3 => {
        r#"void main() {
  var a = {'a','b'};
  var b = {'c','d','e'};
  var c = a.intersection(b);
  print(c.isEmpty);
  print(c.length);
}"#,
        ["true", "0"]
    };

    disjoint_sets_intersection_is_empty_batch_4 => {
        r#"void main() {
  var a = {0,2,4,6,8};
  var b = {1,3,5,7,9};
  var c = a.intersection(b);
  print(c.isEmpty);
  print(c.length);
}"#,
        ["true", "0"]
    };

    contains_all_detects_subset_relation_batch_0 => {
        r#"void main() {
  var sup = {1,2,3,4,5,6,7,8,9,10};
  var sub = {1,2,3,4,5};
  print(sup.containsAll(sub));
}"#,
        ["true"]
    };

    contains_all_detects_subset_relation_batch_1 => {
        r#"void main() {
  var sup = {1,2,3,4,5};
  var sub = {1,2,3,4,5,6,7,8,9,10};
  print(sup.containsAll(sub));
}"#,
        ["false"]
    };

    contains_all_detects_subset_relation_batch_2 => {
        r#"void main() {
  var sup = {1,2,3,4,5};
  var sub = {2,4};
  print(sup.containsAll(sub));
}"#,
        ["true"]
    };

    contains_all_detects_subset_relation_batch_3 => {
        r#"void main() {
  var sup = {1,2,3};
  var sub = {2,4};
  print(sup.containsAll(sub));
}"#,
        ["false"]
    };

    contains_all_detects_subset_relation_batch_4 => {
        r#"void main() {
  var sup = {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20};
  var sub = {5,6,7,8,9,10,11,12,13,14,15};
  print(sup.containsAll(sub));
}"#,
        ["true"]
    };

    contains_all_detects_subset_relation_batch_5 => {
        r#"void main() {
  var sup = {'x','y','z'};
  var sub = {'x'};
  print(sup.containsAll(sub));
}"#,
        ["true"]
    };

    contains_all_detects_subset_relation_batch_6 => {
        r#"void main() {
  var sup = {'y','x'};
  var sub = {'x','y','z'};
  print(sup.containsAll(sub));
}"#,
        ["false"]
    };

    contains_all_detects_subset_relation_batch_7 => {
        r#"void main() {
  var sup = {1,2,3,4,5,6,7,8,9,10};
  var sub = <int>{};
  print(sup.containsAll(sub));
}"#,
        ["true"]
    };

    contains_all_detects_subset_relation_batch_8 => {
        r#"void main() {
  var sup = <int>{};
  var sub = <int>{};
  print(sup.containsAll(sub));
}"#,
        ["true"]
    };

    contains_all_detects_subset_relation_batch_9 => {
        r#"void main() {
  var sup = {1,2,3,4,5};
  var sub = {1,2,3,4,5};
  print(sup.containsAll(sub));
}"#,
        ["true"]
    };

    set_equality_operator_compares_elements_batch_0 => {
        r#"void main() {
  var a = {1,2,3,4,5,6,7,8,9,10};
  var b = {1,2,3,4,5,6,7,8,9,10};
  print(a == b);
}"#,
        ["true"]
    };

    set_equality_operator_compares_elements_batch_1 => {
        r#"void main() {
  var a = {1,2,3,4,5};
  var b = {1,2,3,4,5,6,7,8,9,10};
  print(a == b);
}"#,
        ["false"]
    };

    set_equality_operator_compares_elements_batch_2 => {
        r#"void main() {
  var a = {1,2,3};
  var b = {3,2,1};
  print(a == b);
}"#,
        ["true"]
    };

    set_equality_operator_compares_elements_batch_3 => {
        r#"void main() {
  var a = {1,2,3};
  var b = {1,2,4};
  print(a == b);
}"#,
        ["false"]
    };

    set_equality_operator_compares_elements_batch_4 => {
        r#"void main() {
  var a = {'a','b'};
  var b = {'b','a'};
  print(a == b);
}"#,
        ["true"]
    };

    set_equality_operator_compares_elements_batch_5 => {
        r#"void main() {
  var a = {1,2,3,4,5};
  var b = {2,3,4,5,6};
  print(a == b);
}"#,
        ["false"]
    };

    set_equality_operator_compares_elements_batch_6 => {
        r#"void main() {
  var a = {10,20,30};
  var b = {10,20,30,40};
  print(a == b);
}"#,
        ["false"]
    };

    set_equality_operator_compares_elements_batch_7 => {
        r#"void main() {
  var a = <int>{};
  var b = <int>{};
  print(a == b);
}"#,
        ["true"]
    };

    set_equality_operator_compares_elements_batch_8 => {
        r#"void main() {
  var a = {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15};
  var b = {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15};
  print(a == b);
}"#,
        ["true"]
    };

    set_equality_operator_compares_elements_batch_9 => {
        r#"void main() {
  var a = {5};
  var b = {5,5,5};
  print(a == b);
}"#,
        ["true"]
    };

    union_then_intersection_with_third_set => {
        r#"void main() {
  var a = {1, 2, 3, 4, 5};
  var b = {3, 4, 5, 6, 7};
  var c = {3, 4, 5, 8, 9};
  var r = a.union(b).intersection(c).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["3", "3,4,5"]
    };

    difference_after_union_preserves_left_only => {
        r#"void main() {
  var a = {1, 2, 3};
  var b = {3, 4, 5};
  var u = a.union(b);
  var d = u.difference({4, 5});
  print(d.length);
  print(d.toList()..sort()..join(','));
}"#,
        ["3", "1,2,3"]
    };

    symmetric_gap_via_union_minus_intersection => {
        r#"void main() {
  var a = {1, 2, 3, 4};
  var b = {3, 4, 5, 6};
  var u = a.union(b);
  var i = a.intersection(b);
  var sym = u.difference(i).toList()..sort();
  print(sym.length);
  print(sym.join(','));
}"#,
        ["4", "1,2,5,6"]
    };

    triple_union_accumulates_unique_elements => {
        r#"void main() {
  var a = {1, 2};
  var b = {2, 3};
  var c = {3, 4, 5};
  var r = a.union(b).union(c).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["5", "1,2,3,4,5"]
    };

    intersection_of_three_overlapping_ranges => {
        r#"void main() {
  var a = {for (var i = 1; i < 21; i++) i};
  var b = {for (var i = 5; i < 16; i++) i};
  var c = {for (var i = 8; i < 13; i++) i};
  var r = a.intersection(b).intersection(c).toList()..sort();
  print(r.length);
  print(r.join(','));
}"#,
        ["5", "8,9,10,11,12"]
    };

    mutual_contains_all_implies_equal_large_sets => {
        r#"void main() {
  var a = {for (var i = 0; i < 12; i++) i * 2};
  var b = {0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22};
  print(a.containsAll(b));
  print(b.containsAll(a));
  print(a == b);
}"#,
        ["true", "true", "true"]
    };

    proper_superset_contains_all_but_not_reverse => {
        r#"void main() {
  var big = {for (var i = 1; i <= 15; i++) i};
  var small = {for (var i = 5; i <= 10; i++) i};
  print(big.containsAll(small));
  print(small.containsAll(big));
}"#,
        ["true", "false"]
    };

    union_with_self_is_unchanged_large_set => {
        r#"void main() {
  var s = {for (var i = 1; i <= 20; i++) i};
  var u = s.union(s);
  print(u.length);
  print(u == s);
}"#,
        ["20", "true"]
    };

    intersection_with_self_returns_equal_copy => {
        r#"void main() {
  var s = {10, 20, 30, 40, 50};
  var i = s.intersection(s);
  print(i.length);
  print(i == s);
}"#,
        ["5", "true"]
    };

    difference_with_self_yields_empty_set => {
        r#"void main() {
  var s = {for (var i = 1; i <= 10; i++) i};
  var d = s.difference(s);
  print(d.isEmpty);
  print(d.length);
}"#,
        ["true", "0"]
    };
}
