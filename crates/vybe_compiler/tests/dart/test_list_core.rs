//! List construction, mutation, and Iterable methods.

dart_cases! {
    list_literal_reports_length_three => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.length);
}"#,
        ["3"]
    };

    empty_list_literal_has_zero_length => {
        r#"void main() {
  var list = <int>[];
  print(list.length);
}"#,
        ["0"]
    };

    empty_list_is_empty => {
        r#"void main() {
  var list = <int>[];
  print(list.isEmpty);
}"#,
        ["true"]
    };

    non_empty_list_is_not_empty => {
        r#"void main() {
  var list = [0];
  print(list.isNotEmpty);
}"#,
        ["true"]
    };

    list_index_access_reads_first_element => {
        r#"void main() {
  var list = [10, 20, 30];
  print(list[0]);
}"#,
        ["10"]
    };

    list_index_access_reads_last_element => {
        r#"void main() {
  var list = [10, 20, 30];
  print(list[2]);
}"#,
        ["30"]
    };

    list_index_assignment_mutates_element => {
        r#"void main() {
  var list = [1, 2, 3];
  list[1] = 99;
  print(list[1]);
}"#,
        ["99"]
    };

    list_add_appends_single_element => {
        r#"void main() {
  var list = [1, 2];
  list.add(3);
  print(list.length);
}"#,
        ["3"]
    };

    list_add_all_appends_multiple_elements => {
        r#"void main() {
  var list = [1];
  list.addAll([2, 3, 4]);
  print(list.join(','));
}"#,
        ["1,2,3,4"]
    };

    list_insert_places_element_at_index => {
        r#"void main() {
  var list = [1, 3];
  list.insert(1, 2);
  print(list.join('-'));
}"#,
        ["1-2-3"]
    };

    list_insert_at_start_shifts_elements => {
        r#"void main() {
  var list = [2, 3];
  list.insert(0, 1);
  print(list.first);
}"#,
        ["1"]
    };

    list_remove_by_value_returns_true_when_found => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.remove(2));
}"#,
        ["true"]
    };

    list_remove_by_value_leaves_remaining_elements => {
        r#"void main() {
  var list = [1, 2, 3];
  list.remove(2);
  print(list.join(','));
}"#,
        ["1,3"]
    };

    list_remove_at_index_drops_element => {
        r#"void main() {
  var list = [10, 20, 30];
  list.removeAt(1);
  print(list.join(','));
}"#,
        ["10,30"]
    };

    list_remove_last_pops_trailing_element => {
        r#"void main() {
  var list = [5, 6, 7];
  print(list.removeLast());
}"#,
        ["7"]
    };

    list_remove_range_deletes_span => {
        r#"void main() {
  var list = [1, 2, 3, 4, 5];
  list.removeRange(1, 4);
  print(list.join(','));
}"#,
        ["1,5"]
    };

    list_clear_empties_all_elements => {
        r#"void main() {
  var list = [1, 2, 3];
  list.clear();
  print(list.isEmpty);
}"#,
        ["true"]
    };

    list_first_returns_head_element => {
        r#"void main() {
  var list = [7, 8, 9];
  print(list.first);
}"#,
        ["7"]
    };

    list_last_returns_tail_element => {
        r#"void main() {
  var list = [7, 8, 9];
  print(list.last);
}"#,
        ["9"]
    };

    list_single_requires_exactly_one_element => {
        r#"void main() {
  var list = [42];
  print(list.single);
}"#,
        ["42"]
    };

    list_index_of_finds_first_match => {
        r#"void main() {
  var list = [10, 20, 30];
  print(list.indexOf(20));
}"#,
        ["1"]
    };

    list_index_of_returns_negative_when_missing => {
        r#"void main() {
  var list = [10, 20, 30];
  print(list.indexOf(99));
}"#,
        ["-1"]
    };

    list_last_index_of_finds_last_match => {
        r#"void main() {
  var list = [1, 2, 2, 3];
  print(list.lastIndexOf(2));
}"#,
        ["2"]
    };

    list_contains_detects_present_value => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.contains(2));
}"#,
        ["true"]
    };

    list_contains_rejects_absent_value => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.contains(9));
}"#,
        ["false"]
    };

    list_sublist_extracts_middle_range => {
        r#"void main() {
  var list = [1, 2, 3, 4, 5];
  var sub = list.sublist(1, 4);
  print(sub.join(','));
}"#,
        ["2,3,4"]
    };

    list_sublist_from_index_to_end => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  var sub = list.sublist(2);
  print(sub.join(','));
}"#,
        ["3,4"]
    };

    list_get_range_yields_matching_slice => {
        r#"void main() {
  var list = [5, 6, 7, 8];
  var slice = list.getRange(1, 3);
  print(slice.join(','));
}"#,
        ["6,7"]
    };

    list_join_concatenates_with_separator => {
        r#"void main() {
  var list = ['a', 'b', 'c'];
  print(list.join('-'));
}"#,
        ["a-b-c"]
    };

    list_join_with_empty_separator => {
        r#"void main() {
  var list = ['x', 'y', 'z'];
  print(list.join(''));
}"#,
        ["xyz"]
    };

    list_reversed_iterates_backwards => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.reversed.join(','));
}"#,
        ["3,2,1"]
    };

    list_sort_orders_ascending_by_default => {
        r#"void main() {
  var list = [3, 1, 2];
  list.sort();
  print(list.join(','));
}"#,
        ["1,2,3"]
    };

    list_sort_with_custom_comparator_descending => {
        r#"void main() {
  var list = [3, 1, 2];
  list.sort((a, b) => b.compareTo(a));
  print(list.join(','));
}"#,
        ["3,2,1"]
    };

    list_skip_drops_leading_elements => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  print(list.skip(2).join(','));
}"#,
        ["3,4"]
    };

    list_take_keeps_leading_elements => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  print(list.take(2).join(','));
}"#,
        ["1,2"]
    };

    list_where_filters_by_predicate => {
        r#"void main() {
  var list = [1, 2, 3, 4, 5];
  print(list.where((n) => n % 2 == 0).join(','));
}"#,
        ["2,4"]
    };

    list_map_transforms_each_element => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.map((n) => n * 10).join(','));
}"#,
        ["10,20,30"]
    };

    list_fold_accumulates_with_seed => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  print(list.fold(100, (acc, n) => acc + n));
}"#,
        ["110"]
    };

    list_reduce_combines_without_seed => {
        r#"void main() {
  var list = [2, 3, 4];
  print(list.reduce((a, b) => a * b));
}"#,
        ["24"]
    };

    list_any_detects_matching_element => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.any((n) => n > 2));
}"#,
        ["true"]
    };

    list_every_checks_all_elements_match => {
        r#"void main() {
  var list = [2, 4, 6];
  print(list.every((n) => n % 2 == 0));
}"#,
        ["true"]
    };

    list_element_at_reads_by_position => {
        r#"void main() {
  var list = [10, 20, 30];
  print(list.elementAt(1));
}"#,
        ["20"]
    };

    list_index_where_finds_first_match => {
        r#"void main() {
  var list = [1, 3, 4, 5];
  print(list.indexWhere((n) => n % 2 == 0));
}"#,
        ["2"]
    };

    list_last_index_where_finds_last_match => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  print(list.lastIndexWhere((n) => n % 2 == 0));
}"#,
        ["3"]
    };

    list_fill_range_overwrites_span => {
        r#"void main() {
  var list = [1, 2, 3, 4, 5];
  list.fillRange(1, 4, 0);
  print(list.join(','));
}"#,
        ["1,0,0,0,5"]
    };

    list_set_all_copies_from_source => {
        r#"void main() {
  var list = [0, 0, 0, 0];
  list.setAll(1, [7, 8]);
  print(list.join(','));
}"#,
        ["0,7,8,0"]
    };

    list_remove_where_drops_matching_elements => {
        r#"void main() {
  var list = [1, 2, 3, 4, 5];
  list.removeWhere((n) => n % 2 == 0);
  print(list.join(','));
}"#,
        ["1,3,5"]
    };

    list_retain_where_keeps_matching_elements => {
        r#"void main() {
  var list = [1, 2, 3, 4, 5];
  list.retainWhere((n) => n % 2 == 0);
  print(list.join(','));
}"#,
        ["2,4"]
    };

    list_cast_narrows_element_type => {
        r#"void main() {
  var list = List<Object>.from([1, 2, 3]);
  var nums = list.cast<int>();
  print(nums.reduce((a, b) => a + b));
}"#,
        ["6"]
    };

    list_as_map_pairs_indices_with_values => {
        r#"void main() {
  var list = ['a', 'b', 'c'];
  var map = list.asMap();
  print(map[1]);
}"#,
        ["b"]
    };

    list_followed_by_concatenates_iterables => {
        r#"void main() {
  var first = [1, 2];
  var combined = first.followedBy([3, 4]);
  print(combined.join(','));
}"#,
        ["1,2,3,4"]
    };

    list_expand_flattens_nested_iterables => {
        r#"void main() {
  var nested = [[1, 2], [3], [4, 5]];
  print(nested.expand((part) => part).join(','));
}"#,
        ["1,2,3,4,5"]
    };

    list_to_list_materializes_iterable => {
        r#"void main() {
  Iterable<int> it = [9, 8, 7];
  var copy = it.toList();
  print(copy.join(','));
}"#,
        ["9,8,7"]
    };

    list_single_or_null_returns_element_when_unique => {
        r#"void main() {
  var list = [99];
  print(list.singleOrNull);
}"#,
        ["99"]
    };

    list_single_or_null_returns_null_when_multiple => {
        r#"void main() {
  var list = [1, 2];
  print(list.singleOrNull);
}"#,
        ["null"]
    };

    list_first_where_finds_matching_element => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  print(list.firstWhere((n) => n > 2));
}"#,
        ["3"]
    };

    list_last_where_finds_last_matching_element => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  print(list.lastWhere((n) => n % 2 == 0));
}"#,
        ["4"]
    };

    list_generate_builds_from_index_callback => {
        r#"void main() {
  var list = List.generate(4, (i) => i * i);
  print(list.join(','));
}"#,
        ["0,1,4,9"]
    };

    list_filled_repeats_single_value => {
        r#"void main() {
  var list = List.filled(3, 7);
  print(list.join(','));
}"#,
        ["7,7,7"]
    };

    list_of_copies_from_iterable => {
        r#"void main() {
  var list = List.of([3, 1, 2]);
  list.sort();
  print(list.join(','));
}"#,
        ["1,2,3"]
    };

    list_from_copies_growable_list => {
        r#"void main() {
  var list = List.from([1, 2]);
  list.add(3);
  print(list.length);
}"#,
        ["3"]
    };

    list_set_range_replaces_span => {
        r#"void main() {
  var list = [1, 2, 3, 4, 5];
  list.setRange(1, 3, [8, 9]);
  print(list.join(','));
}"#,
        ["1,8,9,4,5"]
    };

    list_foreach_invokes_callback_per_element => {
        r#"void main() {
  var sum = 0;
  [1, 2, 3].forEach((n) { sum += n; });
  print(sum);
}"#,
        ["6"]
    };

    list_cascade_add_chains_mutations => {
        r#"void main() {
  var list = <int>[];
  list..add(1)..add(2)..add(3);
  print(list.join(','));
}"#,
        ["1,2,3"]
    };
}
