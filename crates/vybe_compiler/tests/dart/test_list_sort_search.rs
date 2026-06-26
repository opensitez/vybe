//! List sort variants and search: indexOf on sorted lists, manual lowerBound.

dart_cases! {
    sort_integers_ascending_default_comparator => {
        r#"void main() {
  var list = [5, 1, 4, 2, 3];
  list.sort();
  print(list.join(','));
}"#,
        ["1,2,3,4,5"]
    };

    sort_already_sorted_list_is_noop_order => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  list.sort();
  print(list.join(','));
}"#,
        ["1,2,3,4"]
    };

    sort_reverse_sorted_list_reorders_ascending => {
        r#"void main() {
  var list = [4, 3, 2, 1];
  list.sort();
  print(list.join(','));
}"#,
        ["1,2,3,4"]
    };

    sort_single_element_list_unchanged => {
        r#"void main() {
  var list = [42];
  list.sort();
  print(list.join(','));
  print(list.length);
}"#,
        ["42", "1"]
    };

    sort_empty_list_remains_empty => {
        r#"void main() {
  var list = <int>[];
  list.sort();
  print(list.isEmpty);
  print(list.length);
}"#,
        ["true", "0"]
    };

    sort_with_duplicates_groups_equal_values => {
        r#"void main() {
  var list = [3, 1, 2, 1, 3];
  list.sort();
  print(list.join(','));
}"#,
        ["1,1,2,3,3"]
    };

    sort_all_equal_elements_unchanged_count => {
        r#"void main() {
  var list = [7, 7, 7, 7];
  list.sort();
  print(list.join(','));
  print(list.length);
}"#,
        ["7,7,7,7", "4"]
    };

    sort_negative_numbers_ascending => {
        r#"void main() {
  var list = [-1, -5, 0, 3, -3];
  list.sort();
  print(list.join(','));
}"#,
        ["-5,-3,-1,0,3"]
    };

    sort_strings_lexicographic_default => {
        r#"void main() {
  var list = ['cherry', 'apple', 'banana'];
  list.sort();
  print(list.join('|'));
}"#,
        ["apple|banana|cherry"]
    };

    sort_strings_case_sensitive_ascii_order => {
        r#"void main() {
  var list = ['b', 'A', 'a', 'B'];
  list.sort();
  print(list.join(','));
}"#,
        ["A,B,a,b"]
    };

    sort_descending_with_custom_comparator => {
        r#"void main() {
  var list = [1, 3, 2, 5, 4];
  list.sort((a, b) => b.compareTo(a));
  print(list.join(','));
}"#,
        ["5,4,3,2,1"]
    };

    sort_by_absolute_value_custom_comparator => {
        r#"void main() {
  var list = [-10, 3, -1, 8, -5];
  list.sort((a, b) => a.abs().compareTo(b.abs()));
  print(list.join(','));
}"#,
        ["-1,3,-5,8,-10"]
    };

    sort_strings_by_length_not_lexicographic => {
        r#"void main() {
  var list = ['aaa', 'b', 'cc', 'd'];
  list.sort((a, b) => a.length.compareTo(b.length));
  print(list.join(','));
}"#,
        ["b,d,cc,aaa"]
    };

    sort_pairs_by_second_component_via_wrapper => {
        r#"void main() {
  var list = ['x:3', 'y:1', 'z:2'];
  list.sort((a, b) {
    var av = int.parse(a.split(':')[1]);
    var bv = int.parse(b.split(':')[1]);
    return av.compareTo(bv);
  });
  print(list.join('|'));
}"#,
        ["y:1|z:2|x:3"]
    };

    sort_then_index_of_finds_existing_target => {
        r#"void main() {
  var list = [30, 10, 20, 40];
  list.sort();
  print(list.indexOf(20));
}"#,
        ["1"]
    };

    sort_then_index_of_returns_negative_for_missing => {
        r#"void main() {
  var list = [3, 1, 2];
  list.sort();
  print(list.indexOf(99));
}"#,
        ["-1"]
    };

    sort_then_index_of_finds_first_duplicate => {
        r#"void main() {
  var list = [2, 1, 2, 3];
  list.sort();
  print(list.indexOf(2));
  print(list.lastIndexOf(2));
}"#,
        ["0", "1"]
    };

    sort_then_binary_search_style_index_of_at_zero => {
        r#"void main() {
  var list = [5, 3, 1, 4, 2];
  list.sort();
  print(list.indexOf(1));
  print(list[0]);
}"#,
        ["0", "1"]
    };

    sort_then_index_of_on_last_element => {
        r#"void main() {
  var list = [1, 5, 3, 4, 2];
  list.sort();
  print(list.indexOf(5));
  print(list.last);
}"#,
        ["4", "5"]
    };

    manual_lower_bound_finds_insertion_index => {
        r#"void main() {
  var list = [1, 3, 5, 7, 9];
  var target = 6;
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  print(lo);
}"#,
        ["3"]
    };

    manual_lower_bound_target_before_all => {
        r#"void main() {
  var list = [2, 4, 6];
  var target = 0;
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  print(lo);
}"#,
        ["0"]
    };

    manual_lower_bound_target_after_all => {
        r#"void main() {
  var list = [2, 4, 6];
  var target = 10;
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  print(lo);
}"#,
        ["3"]
    };

    manual_lower_bound_exact_match_returns_first_index => {
        r#"void main() {
  var list = [1, 2, 2, 2, 3];
  var target = 2;
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  print(lo);
  print(list[lo]);
}"#,
        ["1", "2"]
    };

    manual_lower_bound_on_empty_list => {
        r#"void main() {
  var list = <int>[];
  var target = 5;
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  print(lo);
}"#,
        ["0"]
    };

    manual_upper_bound_style_search => {
        r#"void main() {
  var list = [1, 2, 2, 2, 3];
  var target = 2;
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] <= target) { lo = mid + 1; } else { hi = mid; }
  }
  print(lo);
}"#,
        ["4"]
    };

    sort_then_reversed_join_descending_view => {
        r#"void main() {
  var list = [3, 1, 2];
  list.sort();
  print(list.reversed.join(','));
  print(list.join(','));
}"#,
        ["3,2,1", "1,2,3"]
    };

    sort_descending_then_reversed_yields_ascending => {
        r#"void main() {
  var list = [3, 1, 2];
  list.sort((a, b) => b.compareTo(a));
  print(list.reversed.join(','));
}"#,
        ["1,2,3"]
    };

    sort_stable_relative_order_with_equal_keys => {
        r#"void main() {
  var list = ['b:2', 'a:1', 'c:2', 'd:1'];
  list.sort((a, b) {
    var av = int.parse(a.split(':')[1]);
    var bv = int.parse(b.split(':')[1]);
    return av.compareTo(bv);
  });
  print(list.join('|'));
}"#,
        ["a:1|d:1|b:2|c:2"]
    };

    sort_doubles_ascending => {
        r#"void main() {
  var list = [3.3, 1.1, 2.2];
  list.sort();
  print(list.join(','));
}"#,
        ["1.1,2.2,3.3"]
    };

    sort_mixed_sign_integers_with_zero => {
        r#"void main() {
  var list = [0, -2, 5, -1, 3];
  list.sort();
  print(list.first);
  print(list.last);
}"#,
        ["-2", "5"]
    };

    sort_large_range_integers => {
        r#"void main() {
  var list = [1000, 10, 500, 50];
  list.sort();
  print(list.join(','));
}"#,
        ["10,50,500,1000"]
    };

    sort_two_element_swap => {
        r#"void main() {
  var list = [2, 1];
  list.sort();
  print(list.join(','));
}"#,
        ["1,2"]
    };

    sort_two_element_already_ordered => {
        r#"void main() {
  var list = [1, 2];
  list.sort();
  print(list.join(','));
}"#,
        ["1,2"]
    };

    sort_then_contains_check => {
        r#"void main() {
  var list = [9, 1, 5, 3];
  list.sort();
  print(list.contains(5));
  print(list.contains(8));
}"#,
        ["true", "false"]
    };

    sort_then_first_and_last => {
        r#"void main() {
  var list = [4, 2, 7, 1];
  list.sort();
  print(list.first);
  print(list.last);
}"#,
        ["1", "7"]
    };

    sort_strings_then_index_of_word => {
        r#"void main() {
  var list = ['dog', 'cat', 'ant'];
  list.sort();
  print(list.indexOf('cat'));
  print(list.join(','));
}"#,
        ["1", "ant,cat,dog"]
    };

    sort_with_comparator_using_compare_to_chain => {
        r#"void main() {
  var list = [30, 10, 20];
  list.sort((a, b) => a.compareTo(b));
  print(list.join(','));
}"#,
        ["10,20,30"]
    };

    sort_sublist_copy_leaves_original => {
        r#"void main() {
  var list = [3, 1, 2];
  var sorted = list.toList()..sort();
  print(list.join(','));
  print(sorted.join(','));
}"#,
        ["3,1,2", "1,2,3"]
    };

    sort_then_manual_search_finds_target => {
        r#"void main() {
  var list = [8, 2, 6, 4];
  list.sort();
  var target = 6;
  var found = -1;
  var lo = 0;
  var hi = list.length - 1;
  while (lo <= hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] == target) { found = mid; break; }
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid - 1; }
  }
  print(found);
}"#,
        ["2"]
    };

    sort_then_manual_search_missing_target => {
        r#"void main() {
  var list = [1, 3, 5, 7];
  list.sort();
  var target = 4;
  var found = -1;
  var lo = 0;
  var hi = list.length - 1;
  while (lo <= hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] == target) { found = mid; break; }
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid - 1; }
  }
  print(found);
}"#,
        ["-1"]
    };

    sort_then_index_of_with_start_hint => {
        r#"void main() {
  var list = [1, 2, 2, 2, 3];
  list.sort();
  print(list.indexOf(2, 2));
}"#,
        ["2"]
    };

    sort_then_last_index_of_on_duplicates => {
        r#"void main() {
  var list = [2, 1, 2, 3, 2];
  list.sort();
  print(list.lastIndexOf(2));
}"#,
        ["3"]
    };

    sort_char_codes_as_ints => {
        r#"void main() {
  var list = ['z'.codeUnitAt(0), 'a'.codeUnitAt(0), 'm'.codeUnitAt(0)];
  list.sort();
  print(String.fromCharCodes(list));
}"#,
        ["amz"]
    };

    sort_booleans_false_before_true => {
        r#"void main() {
  var list = [true, false, true, false];
  list.sort((a, b) => (a ? 1 : 0).compareTo(b ? 1 : 0));
  print(list.map((b) => b ? 'T' : 'F').join(''));
}"#,
        ["FFTT"]
    };

    sort_even_odd_partition_via_comparator => {
        r#"void main() {
  var list = [1, 2, 3, 4, 5, 6];
  list.sort((a, b) {
    var ae = a.isEven ? 0 : 1;
    var be = b.isEven ? 0 : 1;
    if (ae != be) return ae.compareTo(be);
    return a.compareTo(b);
  });
  print(list.join(','));
}"#,
        ["2,4,6,1,3,5"]
    };

    sort_three_way_median_first => {
        r#"void main() {
  var list = [2, 3, 1];
  list.sort();
  print(list[1]);
}"#,
        ["2"]
    };

    sort_preserves_length_after_mutation => {
        r#"void main() {
  var list = [5, 4, 3, 2, 1];
  list.sort();
  print(list.length);
  print(list.join('-'));
}"#,
        ["5", "1-2-3-4-5"]
    };

    sort_cascade_on_copy_with_spread => {
        r#"void main() {
  var original = [3, 1, 2];
  var sorted = [...original]..sort();
  print(original.join(','));
  print(sorted.join(','));
}"#,
        ["3,1,2", "1,2,3"]
    };

    sort_then_lower_bound_matches_index_of_for_unique => {
        r#"void main() {
  var list = [9, 3, 7, 1, 5];
  list.sort();
  var target = 7;
  var idx = list.indexOf(target);
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  print(idx);
  print(lo);
  print(list[idx]);
}"#,
        ["3", "3", "7"]
    };

    sort_descending_strings_by_length => {
        r#"void main() {
  var list = ['a', 'bbb', 'cc'];
  list.sort((a, b) => b.length.compareTo(a.length));
  print(list.join(','));
}"#,
        ["bbb,cc,a"]
    };

    sort_integers_then_reversed_list_is_new_view => {
        r#"void main() {
  var list = [1, 3, 2];
  list.sort();
  var rev = list.reversed.toList();
  print(rev.join(','));
  print(list.join(','));
}"#,
        ["3,2,1", "1,2,3"]
    };

    sort_with_identity_comparator_on_sorted_data => {
        r#"void main() {
  var list = [10, 20, 30];
  list.sort((a, b) => a.compareTo(b));
  print(list.indexOf(20));
  print(list.join(','));
}"#,
        ["1", "10,20,30"]
    };

    sort_mixed_duplicate_edges => {
        r#"void main() {
  var list = [1, 2, 1, 3, 2, 1];
  list.sort();
  print(list.where((n) => n == 1).length);
  print(list.first);
  print(list.last);
}"#,
        ["3", "1", "3"]
    };

    sort_then_lower_bound_on_single_element => {
        r#"void main() {
  var list = [5];
  list.sort();
  var target = 5;
  var lo = 0;
  var hi = list.length;
  while (lo < hi) {
    var mid = lo + ((hi - lo) >> 1);
    if (list[mid] < target) { lo = mid + 1; } else { hi = mid; }
  }
  print(lo);
  print(list.indexOf(target));
}"#,
        ["0", "0"]
    };

    sort_strings_ignore_case_via_comparator => {
        r#"void main() {
  var list = ['Banana', 'apple', 'Cherry'];
  list.sort((a, b) => a.toLowerCase().compareTo(b.toLowerCase()));
  print(list.join('|'));
}"#,
        ["apple|Banana|Cherry"]
    };
}
