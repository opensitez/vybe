use crate::helpers::{run_in_main, run_main};

#[test]
fn arrays_sort_orders_descending_input_to_ascending() {
    let out = run_main(
        "int[] nums = {9, 5, 7, 1}; java.util.Arrays.sort(nums); System.out.println(nums[0]); System.out.println(nums[3]);",
    );
    assert_eq!(out, vec!["1", "9"]);
}

#[test]
fn arrays_sort_string_array_lexicographically() {
    let out = run_main(
        "String[] words = {\"cherry\", \"apple\", \"banana\"}; java.util.Arrays.sort(words); System.out.println(words[0]); System.out.println(words[2]);",
    );
    assert_eq!(out, vec!["apple", "cherry"]);
}

#[test]
fn arrays_sort_already_sorted_array_keeps_endpoints() {
    let out = run_main(
        "int[] nums = {2, 4, 6, 8}; java.util.Arrays.sort(nums); System.out.println(nums[0]); System.out.println(nums[3]);",
    );
    assert_eq!(out, vec!["2", "8"]);
}

#[test]
fn arrays_parallel_sort_mention_via_sort_orders_large_sequence() {
    let out = run_main(
        "int[] nums = {8, 1, 6, 3, 9, 2, 7, 4, 5}; java.util.Arrays.sort(nums); System.out.println(nums[0]); System.out.println(nums[8]);",
    );
    assert_eq!(out, vec!["1", "9"]);
}

#[test]
fn arrays_binary_search_finds_existing_middle_index() {
    let out = run_main(
        "int[] nums = {1, 3, 5, 7, 9}; System.out.println(java.util.Arrays.binarySearch(nums, 5));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arrays_binary_search_returns_negative_when_key_absent() {
    let out = run_main(
        "int[] nums = {2, 4, 6, 8}; System.out.println(java.util.Arrays.binarySearch(nums, 5));",
    );
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn arrays_binary_search_finds_first_element_index() {
    let out = run_main(
        "int[] nums = {10, 20, 30}; System.out.println(java.util.Arrays.binarySearch(nums, 10));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arrays_binary_search_finds_last_element_index() {
    let out = run_main(
        "int[] nums = {10, 20, 30}; System.out.println(java.util.Arrays.binarySearch(nums, 30));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arrays_binary_search_after_sort_finds_reinserted_value() {
    let out = run_main(
        "int[] nums = {9, 1, 5}; java.util.Arrays.sort(nums); System.out.println(java.util.Arrays.binarySearch(nums, 5));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn arrays_equals_true_for_identical_int_arrays() {
    let out = run_main(
        "int[] a = {1, 2, 3}; int[] b = {1, 2, 3}; System.out.println(java.util.Arrays.equals(a, b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_equals_false_when_elements_differ() {
    let out = run_main(
        "int[] a = {1, 2, 3}; int[] b = {1, 9, 3}; System.out.println(java.util.Arrays.equals(a, b));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arrays_equals_false_for_different_lengths() {
    let out = run_main(
        "int[] a = {1, 2}; int[] b = {1, 2, 3}; System.out.println(java.util.Arrays.equals(a, b));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arrays_equals_true_for_same_array_reference() {
    let out = run_main(
        "int[] nums = {1, 2}; System.out.println(java.util.Arrays.equals(nums, nums));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_deep_equals_true_for_matching_nested_int_arrays() {
    let out = run_main(
        "int[][] left = {{1, 2}, {3, 4}}; int[][] right = {{1, 2}, {3, 4}}; System.out.println(java.util.Arrays.deepEquals(left, right));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_deep_equals_false_when_inner_row_differs() {
    let out = run_main(
        "int[][] left = {{1, 2}, {3, 4}}; int[][] right = {{1, 2}, {3, 9}}; System.out.println(java.util.Arrays.deepEquals(left, right));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arrays_deep_equals_true_for_empty_nested_arrays() {
    let out = run_main(
        "int[][] left = {}; int[][] right = {}; System.out.println(java.util.Arrays.deepEquals(left, right));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_fill_entire_int_array_with_constant() {
    let out = run_main(
        "int[] nums = new int[4]; java.util.Arrays.fill(nums, 9); System.out.println(nums[0]); System.out.println(nums[3]);",
    );
    assert_eq!(out, vec!["9", "9"]);
}

#[test]
fn arrays_fill_partial_range_leaves_outside_indices() {
    let out = run_main(
        "int[] nums = {1, 1, 1, 1, 1}; java.util.Arrays.fill(nums, 1, 4, 7); System.out.println(nums[0]); System.out.println(nums[2]); System.out.println(nums[4]);",
    );
    assert_eq!(out, vec!["1", "7", "1"]);
}

#[test]
fn arrays_fill_on_new_array_overwrites_default_zeros() {
    let out = run_main(
        "int[] nums = new int[3]; java.util.Arrays.fill(nums, 4); System.out.println(nums[1]);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn arrays_copy_of_extends_tail_with_zero_padding() {
    let out = run_main(
        "int[] src = {1, 2, 3}; int[] copy = java.util.Arrays.copyOf(src, 5); System.out.println(copy.length); System.out.println(copy[4]);",
    );
    assert_eq!(out, vec!["5", "0"]);
}

#[test]
fn arrays_copy_of_truncates_to_shorter_length() {
    let out = run_main(
        "int[] src = {1, 2, 3, 4}; int[] copy = java.util.Arrays.copyOf(src, 2); System.out.println(copy.length); System.out.println(copy[1]);",
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn arrays_copy_of_zero_length_yields_empty_array() {
    let out = run_main(
        "int[] src = {1, 2}; int[] copy = java.util.Arrays.copyOf(src, 0); System.out.println(copy.length);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arrays_copy_of_range_extracts_middle_slice() {
    let out = run_main(
        "int[] src = {1, 2, 3, 4, 5}; int[] mid = java.util.Arrays.copyOfRange(src, 1, 4); System.out.println(mid.length); System.out.println(mid[0]); System.out.println(mid[2]);",
    );
    assert_eq!(out, vec!["3", "2", "4"]);
}

#[test]
fn arrays_copy_of_range_single_element_slice() {
    let out = run_main(
        "int[] src = {10, 20, 30}; int[] slice = java.util.Arrays.copyOfRange(src, 1, 2); System.out.println(slice.length); System.out.println(slice[0]);",
    );
    assert_eq!(out, vec!["1", "20"]);
}

#[test]
fn arrays_as_list_wraps_array_literal_elements() {
    let out = run_main(
        "java.util.List<Integer> list = java.util.Arrays.asList(5, 6, 7); System.out.println(list.get(1)); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["6", "3"]);
}

#[test]
fn arrays_as_list_from_int_array_variable_preserves_order() {
    let out = run_main(
        "Integer[] data = {1, 2, 3}; java.util.List<Integer> list = java.util.Arrays.asList(data); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn arrays_as_list_get_reflects_underlying_array_mutation() {
    let out = run_main(
        "int[] data = {1, 2, 3}; java.util.List<Integer> list = java.util.Arrays.asList(data); data[1] = 99; System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn arrays_stream_count_matches_array_length() {
    let out = run_main(
        "int[] nums = {1, 2, 3, 4}; System.out.println(java.util.Arrays.stream(nums).count());",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn arrays_stream_map_doubles_each_element_via_collect() {
    let out = run_main(
        "int[] nums = {1, 2, 3}; java.util.List<Integer> doubled = java.util.Arrays.stream(nums).map(n -> n * 2).toList(); System.out.println(doubled.get(0)); System.out.println(doubled.get(2));",
    );
    assert_eq!(out, vec!["2", "6"]);
}

#[test]
fn arrays_stream_filter_reduces_element_count() {
    let out = run_main(
        "int[] nums = {1, 2, 3, 4, 5}; long count = java.util.Arrays.stream(nums).filter(n -> n % 2 == 0).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arrays_to_string_formats_int_array_with_brackets() {
    let out = run_main(
        "int[] nums = {1, 2, 3}; System.out.println(java.util.Arrays.toString(nums));",
    );
    assert_eq!(out, vec!["[1, 2, 3]"]);
}

#[test]
fn arrays_to_string_on_empty_array_is_empty_brackets() {
    let out = run_main(
        "int[] nums = {}; System.out.println(java.util.Arrays.toString(nums));",
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn arrays_to_string_single_element_array() {
    let out = run_main(
        "int[] nums = {42}; System.out.println(java.util.Arrays.toString(nums));",
    );
    assert_eq!(out, vec!["[42]"]);
}

#[test]
fn arrays_compare_returns_zero_for_equal_arrays() {
    let out = run_main(
        "int[] a = {1, 2, 3}; int[] b = {1, 2, 3}; System.out.println(java.util.Arrays.compare(a, b));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arrays_compare_negative_when_first_array_lexicographically_smaller() {
    let out = run_main(
        "int[] a = {1, 2, 3}; int[] b = {1, 3, 3}; System.out.println(java.util.Arrays.compare(a, b));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn arrays_compare_positive_when_first_array_lexicographically_larger() {
    let out = run_main(
        "int[] a = {5, 1}; int[] b = {2, 9}; System.out.println(java.util.Arrays.compare(a, b));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn arrays_compare_single_element_equal_arrays() {
    let out = run_main(
        "int[] a = {7}; int[] b = {7}; System.out.println(java.util.Arrays.compare(a, b));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arrays_mismatch_negative_one_when_arrays_equal() {
    let out = run_main(
        "int[] a = {1, 2, 3}; int[] b = {1, 2, 3}; System.out.println(java.util.Arrays.mismatch(a, b));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn arrays_mismatch_index_when_first_difference_found() {
    let out = run_main(
        "int[] a = {1, 2, 3}; int[] b = {1, 9, 3}; System.out.println(java.util.Arrays.mismatch(a, b));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn arrays_mismatch_on_different_length_arrays() {
    let out = run_main(
        "int[] a = {1, 2}; int[] b = {1, 2, 3}; System.out.println(java.util.Arrays.mismatch(a, b));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arrays_set_all_fills_via_index_mapper_lambda() {
    let out = run_main(
        "int[] nums = new int[4]; java.util.Arrays.setAll(nums, i -> i * 10); System.out.println(nums[0]); System.out.println(nums[3]);",
    );
    assert_eq!(out, vec!["0", "30"]);
}

#[test]
fn arrays_set_all_on_zero_length_array_is_noop() {
    let out = run_main(
        "int[] nums = {}; java.util.Arrays.setAll(nums, i -> i); System.out.println(nums.length);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arrays_fill_then_equals_detects_matching_copy() {
    let out = run_main(
        "int[] nums = new int[3]; java.util.Arrays.fill(nums, 4); int[] copy = java.util.Arrays.copyOf(nums, 3); System.out.println(java.util.Arrays.equals(nums, copy));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_copy_of_same_length_is_element_wise_equal() {
    let out = run_main(
        "int[] src = {8, 9}; int[] copy = java.util.Arrays.copyOf(src, src.length); System.out.println(java.util.Arrays.equals(src, copy));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_deep_equals_detects_row_count_difference() {
    let types = r#"
        static boolean rowsMatch(int[][] a, int[][] b) {
            return java.util.Arrays.deepEquals(a, b);
        }
    "#;
    let out = run_in_main(
        "int[][] left = {{1}}; int[][] right = {{1}, {2}}; System.out.println(rowsMatch(left, right));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}
