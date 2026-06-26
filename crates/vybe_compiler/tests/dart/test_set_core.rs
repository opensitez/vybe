//! Core Set behaviors: literals, mutation, set algebra, factories, and Iterable methods.

dart_cases! {
    set_literal_reports_length_three => {
        r#"void main() {
  var s = {1, 2, 3};
  print(s.length);
}"#,
        ["3"]
    };

    set_literal_deduplicates_duplicate_entries => {
        r#"void main() {
  var s = {1, 2, 2, 3, 3};
  print(s.length);
}"#,
        ["3"]
    };

    set_typed_empty_literal_has_zero_length => {
        r#"void main() {
  var s = <int>{};
  print(s.length);
  print(s.isEmpty);
}"#,
        ["0", "true"]
    };

    set_empty_is_empty_returns_true => {
        r#"void main() {
  var s = <String>{};
  print(s.isEmpty);
}"#,
        ["true"]
    };

    set_nonempty_is_not_empty => {
        r#"void main() {
  var s = {0};
  print(s.isNotEmpty);
}"#,
        ["true"]
    };

    set_add_inserts_new_element_returns_true => {
        r#"void main() {
  var s = <int>{};
  print(s.add(10));
}"#,
        ["true"]
    };

    set_add_duplicate_element_returns_false => {
        r#"void main() {
  var s = {7};
  print(s.add(7));
  print(s.length);
}"#,
        ["false", "1"]
    };

    set_add_increments_length => {
        r#"void main() {
  var s = <int>{};
  s.add(1);
  s.add(2);
  s.add(3);
  print(s.length);
}"#,
        ["3"]
    };

    set_add_all_merges_new_elements => {
        r#"void main() {
  var s = {1};
  s.addAll({2, 3, 4});
  print(s.length);
  print(s.contains(4));
}"#,
        ["4", "true"]
    };

    set_add_all_skips_already_present => {
        r#"void main() {
  var s = {1, 2};
  s.addAll({2, 3});
  print(s.length);
  print(s.contains(3));
}"#,
        ["3", "true"]
    };

    set_remove_existing_element_returns_true => {
        r#"void main() {
  var s = {5, 6, 7};
  print(s.remove(6));
}"#,
        ["true"]
    };

    set_remove_missing_element_returns_false => {
        r#"void main() {
  var s = {1, 2};
  print(s.remove(99));
}"#,
        ["false"]
    };

    set_remove_decrements_length => {
        r#"void main() {
  var s = {10, 20, 30};
  s.remove(20);
  print(s.length);
  print(s.contains(20));
}"#,
        ["2", "false"]
    };

    set_remove_all_deletes_matching_elements => {
        r#"void main() {
  var s = {1, 2, 3, 4};
  s.removeAll({2, 4});
  print(s.toList().join(','));
  print(s.length);
}"#,
        ["1,3", "2"]
    };

    set_retain_all_keeps_only_shared_elements => {
        r#"void main() {
  var s = {1, 2, 3, 4};
  s.retainAll({2, 3, 9});
  print(s.toList().join(','));
  print(s.length);
}"#,
        ["2,3", "2"]
    };

    set_contains_true_when_element_present => {
        r#"void main() {
  var s = {'alpha', 'beta'};
  print(s.contains('alpha'));
}"#,
        ["true"]
    };

    set_contains_false_when_element_absent => {
        r#"void main() {
  var s = {1, 2, 3};
  print(s.contains(0));
}"#,
        ["false"]
    };

    set_lookup_returns_equal_element => {
        r#"void main() {
  var s = {10, 20, 30};
  print(s.lookup(20));
}"#,
        ["20"]
    };

    set_lookup_returns_null_when_absent => {
        r#"void main() {
  var s = {1, 2};
  print(s.lookup(5));
}"#,
        ["null"]
    };

    set_union_combines_distinct_elements => {
        r#"void main() {
  var a = {1, 2};
  var b = {2, 3};
  var c = a.union(b);
  print(c.length);
  print(c.toList().join(','));
}"#,
        ["3", "1,2,3"]
    };

    set_intersection_keeps_common_elements => {
        r#"void main() {
  var a = {1, 2, 3};
  var b = {2, 3, 4};
  print(a.intersection(b).toList().join(','));
  print(a.intersection(b).length);
}"#,
        ["2,3", "2"]
    };

    set_difference_removes_other_elements => {
        r#"void main() {
  var a = {1, 2, 3};
  var b = {2};
  print(a.difference(b).toList().join(','));
  print(a.difference(b).length);
}"#,
        ["1,3", "2"]
    };

    set_from_deduplicates_source_iterable => {
        r#"void main() {
  var s = Set<int>.from([1, 2, 2, 3, 3, 3]);
  print(s.length);
  print(s.contains(2));
}"#,
        ["3", "true"]
    };

    set_of_copies_from_existing_set => {
        r#"void main() {
  var src = {5, 6, 7};
  var copy = Set<int>.of(src);
  print(copy.length);
  print(copy.contains(6));
}"#,
        ["3", "true"]
    };

    set_identity_starts_empty_mutable => {
        r#"void main() {
  var s = Set<int>.identity();
  s.add(42);
  print(s.length);
  print(s.contains(42));
}"#,
        ["1", "true"]
    };

    set_foreach_accumulates_values => {
        r#"void main() {
  var s = {1, 2, 3};
  var sum = 0;
  s.forEach((e) => sum += e);
  print(sum);
}"#,
        ["6"]
    };

    set_map_transforms_each_element => {
        r#"void main() {
  var s = {1, 2, 3};
  var doubled = s.map((e) => e * 2).toList()..sort();
  print(doubled.join(','));
}"#,
        ["2,4,6"]
    };

    set_where_filters_by_predicate => {
        r#"void main() {
  var s = {1, 2, 3, 4, 5};
  var evens = s.where((e) => e % 2 == 0).toList()..sort();
  print(evens.join(','));
}"#,
        ["2,4"]
    };

    set_to_list_materializes_elements => {
        r#"void main() {
  var s = {9, 8, 7};
  var list = s.toList()..sort();
  print(list.join(','));
  print(list.length);
}"#,
        ["7,8,9", "3"]
    };

    set_clear_removes_all_elements => {
        r#"void main() {
  var s = {1, 2, 3};
  s.clear();
  print(s.length);
  print(s.isEmpty);
}"#,
        ["0", "true"]
    };

    set_union_with_empty_equals_original => {
        r#"void main() {
  var a = {1, 2};
  var c = a.union(<int>{});
  print(c.length);
  print(c.contains(1));
  print(c.contains(2));
}"#,
        ["2", "true", "true"]
    };

    set_intersection_disjoint_yields_empty => {
        r#"void main() {
  var a = {1, 2};
  var b = {3, 4};
  var c = a.intersection(b);
  print(c.isEmpty);
  print(c.length);
}"#,
        ["true", "0"]
    };

    set_difference_with_superset => {
        r#"void main() {
  var a = {1, 2, 3};
  var b = {1, 2, 3, 4, 5};
  print(a.difference(b).isEmpty);
  print(a.difference(b).length);
}"#,
        ["true", "0"]
    };

    set_retain_all_on_disjoint_empties_set => {
        r#"void main() {
  var s = {1, 2, 3};
  s.retainAll({4, 5});
  print(s.isEmpty);
  print(s.length);
}"#,
        ["true", "0"]
    };

    set_add_all_from_empty_is_noop => {
        r#"void main() {
  var s = {1, 2};
  s.addAll(<int>{});
  print(s.length);
  print(s.contains(1));
}"#,
        ["2", "true"]
    };

    set_remove_all_on_empty_other_is_noop => {
        r#"void main() {
  var s = {10, 20};
  s.removeAll(<int>{});
  print(s.length);
  print(s.toList().join(','));
}"#,
        ["2", "10,20"]
    };

    set_string_elements_contains_check => {
        r#"void main() {
  var s = {'dart', 'set', 'core'};
  print(s.contains('set'));
  print(s.length);
}"#,
        ["true", "3"]
    };

    set_foreach_prints_insertion_order => {
        r#"void main() {
  var s = {3, 1, 2};
  s.forEach((e) => print(e));
}"#,
        ["3", "1", "2"]
    };

    set_cast_narrows_element_type => {
        r#"void main() {
  var s = Set<Object>.from({1, 2, 3});
  var nums = s.cast<int>();
  print(nums.reduce((a, b) => a + b));
}"#,
        ["6"]
    };

    set_remove_where_drops_matching_elements => {
        r#"void main() {
  var s = {1, 2, 3, 4, 5};
  s.removeWhere((e) => e % 2 == 0);
  print(s.toList().join(','));
  print(s.length);
}"#,
        ["1,3,5", "3"]
    };

    set_retain_where_keeps_matching_elements => {
        r#"void main() {
  var s = {1, 2, 3, 4, 5};
  s.retainWhere((e) => e > 3);
  print(s.toList().join(','));
  print(s.length);
}"#,
        ["4,5", "2"]
    };

    set_any_detects_matching_element => {
        r#"void main() {
  var s = {1, 2, 3};
  print(s.any((e) => e > 2));
}"#,
        ["true"]
    };

    set_every_verifies_all_elements_match => {
        r#"void main() {
  var s = {2, 4, 6};
  print(s.every((e) => e % 2 == 0));
}"#,
        ["true"]
    };

    set_fold_accumulates_without_mutation => {
        r#"void main() {
  var s = {1, 2, 3};
  print(s.fold(10, (acc, e) => acc + e));
  print(s.length);
}"#,
        ["16", "3"]
    };

    set_first_returns_earliest_inserted_element => {
        r#"void main() {
  var s = {30, 10, 20};
  print(s.first);
  print(s.last);
}"#,
        ["30", "20"]
    };
}
