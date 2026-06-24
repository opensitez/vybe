use crate::helpers::{run_in_main, run_main};

#[test]
fn array_literal_reports_length_field() {
    let out = run_main("int[] nums = {10, 20, 30}; System.out.println(nums.length);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn array_literal_reads_first_element_by_index() {
    let out = run_main("int[] nums = {7, 8, 9}; System.out.println(nums[0]);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn array_literal_reads_last_element_by_index() {
    let out = run_main("int[] nums = {7, 8, 9}; System.out.println(nums[2]);");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn new_int_array_zero_fills_all_slots() {
    let out = run_main("int[] nums = new int[4]; System.out.println(nums[0]); System.out.println(nums[3]);");
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn new_int_array_length_matches_allocation_size() {
    let out = run_main("int[] nums = new int[5]; System.out.println(nums.length);");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn array_index_write_persists_for_later_reads() {
    let out = run_main("int[] nums = {1, 2, 3}; nums[1] = 42; System.out.println(nums[1]);");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn arrays_copy_of_extends_tail_with_default_zeros() {
    let out = run_main("int[] src = {1, 2, 3}; int[] copy = java.util.Arrays.copyOf(src, 5); System.out.println(copy.length); System.out.println(copy[4]);");
    assert_eq!(out, vec!["5", "0"]);
}

#[test]
fn arrays_copy_of_truncates_to_shorter_length() {
    let out = run_main("int[] src = {1, 2, 3, 4}; int[] copy = java.util.Arrays.copyOf(src, 2); System.out.println(copy.length); System.out.println(copy[1]);");
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn arrays_copy_of_preserves_existing_elements() {
    let out = run_main("int[] src = {5, 6, 7}; int[] copy = java.util.Arrays.copyOf(src, 3); System.out.println(copy[0]); System.out.println(copy[2]);");
    assert_eq!(out, vec!["5", "7"]);
}

#[test]
fn arrays_copy_of_range_extracts_middle_slice() {
    let out = run_main("int[] src = {1, 2, 3, 4, 5}; int[] mid = java.util.Arrays.copyOfRange(src, 1, 4); System.out.println(mid.length); System.out.println(mid[0]); System.out.println(mid[2]);");
    assert_eq!(out, vec!["3", "2", "4"]);
}

#[test]
fn arrays_fill_sets_every_index_to_value() {
    let out = run_main("int[] nums = new int[4]; java.util.Arrays.fill(nums, 9); System.out.println(nums[0]); System.out.println(nums[3]);");
    assert_eq!(out, vec!["9", "9"]);
}

#[test]
fn arrays_fill_partial_range_leaves_outside_values() {
    let out = run_main("int[] nums = {1, 1, 1, 1, 1}; java.util.Arrays.fill(nums, 1, 4, 7); System.out.println(nums[0]); System.out.println(nums[2]); System.out.println(nums[4]);");
    assert_eq!(out, vec!["1", "7", "1"]);
}

#[test]
fn arrays_sort_orders_integers_ascending() {
    let out = run_main("int[] nums = {3, 1, 4, 1, 5}; java.util.Arrays.sort(nums); System.out.println(nums[0]); System.out.println(nums[4]);");
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn arrays_sort_leaves_already_sorted_array_unchanged() {
    let out = run_main("int[] nums = {2, 4, 6}; java.util.Arrays.sort(nums); System.out.println(nums[1]);");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn arrays_sort_handles_single_element_array() {
    let out = run_main("int[] nums = {42}; java.util.Arrays.sort(nums); System.out.println(nums[0]);");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn arrays_binary_search_finds_existing_value_index() {
    let out = run_main("int[] nums = {1, 3, 5, 7}; System.out.println(java.util.Arrays.binarySearch(nums, 5));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn arrays_binary_search_returns_negative_for_missing_value() {
    let out = run_main("int[] nums = {2, 4, 6, 8}; System.out.println(java.util.Arrays.binarySearch(nums, 5));");
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn arrays_binary_search_on_sorted_after_sort_call() {
    let out = run_main("int[] nums = {9, 1, 5}; java.util.Arrays.sort(nums); System.out.println(java.util.Arrays.binarySearch(nums, 5));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn arrays_equals_true_for_matching_contents() {
    let out = run_main("int[] a = {1, 2, 3}; int[] b = {1, 2, 3}; System.out.println(java.util.Arrays.equals(a, b));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_equals_false_for_different_lengths() {
    let out = run_main("int[] a = {1, 2}; int[] b = {1, 2, 3}; System.out.println(java.util.Arrays.equals(a, b));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arrays_equals_false_for_different_values() {
    let out = run_main("int[] a = {1, 2, 3}; int[] b = {1, 9, 3}; System.out.println(java.util.Arrays.equals(a, b));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn arrays_to_string_formats_int_array() {
    let out = run_main("int[] nums = {1, 2, 3}; System.out.println(java.util.Arrays.toString(nums));");
    assert_eq!(out, vec!["[1, 2, 3]"]);
}

#[test]
fn arrays_to_string_on_empty_array() {
    let out = run_main("int[] nums = {}; System.out.println(java.util.Arrays.toString(nums));");
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn two_dimensional_array_row_count_from_literal() {
    let out = run_main("int[][] grid = {{1, 2}, {3, 4}, {5, 6}}; System.out.println(grid.length);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn two_dimensional_array_column_count_on_first_row() {
    let out = run_main("int[][] grid = {{1, 2, 3}, {4, 5, 6}}; System.out.println(grid[0].length);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn two_dimensional_array_reads_inner_element() {
    let out = run_main("int[][] grid = {{1, 2}, {3, 4}}; System.out.println(grid[1][0]);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn two_dimensional_array_writes_inner_element() {
    let out = run_main("int[][] grid = {{1, 2}, {3, 4}}; grid[0][1] = 99; System.out.println(grid[0][1]);");
    assert_eq!(out, vec!["99"]);
}

#[test]
fn jagged_two_d_array_rows_can_have_different_lengths() {
    let out = run_main("int[][] grid = {{1}, {2, 3, 4}}; System.out.println(grid[0].length); System.out.println(grid[1].length);");
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn nested_int_matrix_deep_equality_via_row_arrays_equals() {
    let types = r#"
        static boolean deepEquals2D(int[][] a, int[][] b) {
            if (a.length != b.length) return false;
            for (int i = 0; i < a.length; i++) {
                if (!java.util.Arrays.equals(a[i], b[i])) return false;
            }
            return true;
        }
    "#;
    let out = run_in_main(
        "int[][] left = {{1, 2}, {3, 4}}; int[][] right = {{1, 2}, {3, 4}}; System.out.println(deepEquals2D(left, right));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn nested_int_matrix_deep_compare_detects_row_difference() {
    let types = r#"
        static boolean deepEquals2D(int[][] a, int[][] b) {
            if (a.length != b.length) return false;
            for (int i = 0; i < a.length; i++) {
                if (!java.util.Arrays.equals(a[i], b[i])) return false;
            }
            return true;
        }
    "#;
    let out = run_in_main(
        "int[][] left = {{1, 2}, {3, 4}}; int[][] right = {{1, 2}, {3, 9}}; System.out.println(deepEquals2D(left, right));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn new_boolean_array_defaults_to_false() {
    let out = run_main("boolean[] flags = new boolean[2]; System.out.println(flags[0]); System.out.println(flags[1]);");
    assert_eq!(out, vec!["false", "false"]);
}

#[test]
fn string_array_initializer_stores_literals() {
    let out = run_main("String[] words = {\"a\", \"bb\"}; System.out.println(words[0]); System.out.println(words[1]);");
    assert_eq!(out, vec!["a", "bb"]);
}

#[test]
fn arrays_copy_of_zero_length_produces_empty_array() {
    let out = run_main("int[] src = {1, 2}; int[] copy = java.util.Arrays.copyOf(src, 0); System.out.println(copy.length);");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn fill_then_copy_of_preserves_filled_values() {
    let out = run_main("int[] nums = new int[3]; java.util.Arrays.fill(nums, 4); int[] copy = java.util.Arrays.copyOf(nums, 3); System.out.println(copy[1]);");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn arrays_binary_search_finds_first_element() {
    let out = run_main("int[] nums = {10, 20, 30}; System.out.println(java.util.Arrays.binarySearch(nums, 10));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arrays_equals_true_for_same_reference() {
    let out = run_main("int[] nums = {1, 2}; System.out.println(java.util.Arrays.equals(nums, nums));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn two_d_literal_nested_loop_prints_all_cells() {
    let out = run_main("int[][] grid = {{1, 2}, {3, 4}}; for (int r = 0; r < grid.length; r++) { for (int c = 0; c < grid[r].length; c++) { System.out.println(grid[r][c]); } }");
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

#[test]
fn copy_of_same_length_is_element_wise_equal() {
    let out = run_main("int[] src = {8, 9}; int[] copy = java.util.Arrays.copyOf(src, src.length); System.out.println(java.util.Arrays.equals(src, copy));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_array_sum_via_indexed_while_loop() {
    let out = run_main("int[] nums = {1, 2, 3, 4}; int i = 0; int sum = 0; while (i < nums.length) { sum += nums[i]; i++; } System.out.println(sum);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn new_two_d_int_array_allocates_rectangular_grid() {
    let out = run_main("int[][] grid = new int[2][3]; System.out.println(grid.length); System.out.println(grid[1].length); System.out.println(grid[0][2]);");
    assert_eq!(out, vec!["2", "3", "0"]);
}
