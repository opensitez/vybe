use crate::helpers::run_main;

#[test]
fn array_initializer_length_field() {
    let out = run_main("int[] nums = {10, 20, 30}; System.out.println(nums.length);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn array_element_read_by_index() {
    let out = run_main("int[] nums = {5, 9, 2}; System.out.println(nums[1]);");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn array_element_assignment_mutates_slot() {
    let out = run_main("int[] nums = {1, 2, 3}; nums[0] = 99; System.out.println(nums[0]);");
    assert_eq!(out, vec!["99"]);
}

#[test]
fn multidimensional_array_row_count() {
    let out = run_main("int[][] grid = {{1, 2}, {3, 4}, {5, 6}}; System.out.println(grid.length);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn array_created_with_new_has_zero_initial_values() {
    let out = run_main("int[] nums = new int[3]; System.out.println(nums[0]); System.out.println(nums[2]);");
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn arrays_copy_of_truncates_or_extends() {
    let out = run_main(
        "int[] src = {1, 2, 3}; int[] copy = java.util.Arrays.copyOf(src, 5); System.out.println(copy.length); System.out.println(copy[4]);",
    );
    assert_eq!(out, vec!["5", "0"]);
}
