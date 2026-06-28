use crate::helpers::run_python_one;

#[test]
fn slice_assign_replace_middle_segment() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\nxs[1:3] = [20, 30]\nprint(xs)\n"),
        "[1, 20, 30, 4]"
    );
}

#[test]
fn slice_assign_insert_without_removing() {
    assert_eq!(
        run_python_one("xs = [1, 3]\nxs[1:1] = [2]\nprint(xs)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn slice_assign_delete_range_via_empty_list() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\nxs[1:3] = []\nprint(xs)\n"),
        "[1, 4]"
    );
}

#[test]
fn slice_assign_full_slice_replacement() {
    assert_eq!(
        run_python_one("xs = [9, 8, 7]\nxs[:] = [1, 2]\nprint(xs)\n"),
        "[1, 2]"
    );
}

#[test]
fn slice_assign_expand_with_longer_iterable() {
    assert_eq!(
        run_python_one("xs = [1, 2]\nxs[1:1] = [10, 11, 12]\nprint(xs)\n"),
        "[1, 10, 11, 12, 2]"
    );
}

#[test]
fn slice_assign_shrink_with_shorter_iterable() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4, 5]\nxs[1:4] = [9]\nprint(xs)\n"),
        "[1, 9, 5]"
    );
}

#[test]
fn slice_assign_from_start_only() {
    assert_eq!(
        run_python_one("xs = [0, 1, 2, 3]\nxs[:2] = [7, 8]\nprint(xs)\n"),
        "[7, 8, 2, 3]"
    );
}

#[test]
fn slice_assign_to_end_only() {
    assert_eq!(
        run_python_one("xs = [0, 1, 2, 3]\nxs[2:] = [20, 30]\nprint(xs)\n"),
        "[0, 1, 20, 30]"
    );
}

#[test]
fn slice_assign_with_step_not_supported_use_plain_slice() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\nxs[0:2] = ['a', 'b']\nprint(xs)\n"),
        "['a', 'b', 3, 4]"
    );
}

#[test]
fn slice_assign_negative_start_index() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\nxs[-2:] = [30, 40]\nprint(xs)\n"),
        "[1, 2, 30, 40]"
    );
}

#[test]
fn slice_assign_negative_stop_index() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\nxs[:-1] = [10]\nprint(xs)\n"),
        "[10, 4]"
    );
}

#[test]
fn del_slice_removes_interior() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\ndel xs[1:3]\nprint(xs)\n"),
        "[1, 4]"
    );
}

#[test]
fn del_slice_full_clears_list() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\ndel xs[:]\nprint(xs)\n"),
        "[]"
    );
}

#[test]
fn del_slice_from_start() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\ndel xs[:2]\nprint(xs)\n"),
        "[3, 4]"
    );
}

#[test]
fn del_slice_to_end() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\ndel xs[2:]\nprint(xs)\n"),
        "[1, 2]"
    );
}

#[test]
fn slice_assign_repeat_element_via_multiplier() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs[1:2] = [0] * 3\nprint(xs)\n"),
        "[1, 0, 0, 0, 3]"
    );
}

#[test]
fn slice_assign_nested_list_literal() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs[1:2] = [[9]]\nprint(xs)\n"),
        "[1, [9], 3]"
    );
}

#[test]
fn slice_assign_after_append_preserves_tail() {
    assert_eq!(
        run_python_one("xs = [1, 2]\nxs.append(3)\nxs[1:2] = [20]\nprint(xs)\n"),
        "[1, 20, 3]"
    );
}

#[test]
fn slice_assign_zero_width_at_end_appends() {
    assert_eq!(
        run_python_one("xs = [1, 2]\nxs[2:2] = [3]\nprint(xs)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn slice_assign_zero_width_at_start_prepends() {
    assert_eq!(
        run_python_one("xs = [2, 3]\nxs[0:0] = [1]\nprint(xs)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn slice_read_after_assign_consistent() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\nxs[1:3] = [8, 9]\nprint(xs[1:3])\n"),
        "[8, 9]"
    );
}

#[test]
fn slice_assign_string_elements_in_list() {
    assert_eq!(
        run_python_one("xs = ['a', 'b', 'c']\nxs[0:1] = ['z']\nprint(xs)\n"),
        "['z', 'b', 'c']"
    );
}

#[test]
fn slice_assign_bool_elements() {
    assert_eq!(
        run_python_one("xs = [True, False, True]\nxs[1:2] = [True]\nprint(xs)\n"),
        "[True, True, True]"
    );
}

#[test]
fn slice_assign_mixed_types_in_replacement() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs[1:2] = ['x', None]\nprint(xs)\n"),
        "[1, 'x', None, 3]"
    );
}

#[test]
fn slice_assign_preserves_unrelated_indices() {
    assert_eq!(
        run_python_one("xs = [10, 20, 30, 40]\nxs[1:3] = [2, 3]\nprint(xs[0], xs[-1])\n"),
        "10 40"
    );
}

#[test]
fn slice_assign_in_function_mutates_caller_list() {
    assert_eq!(
        run_python_one("def patch(xs):\n xs[1:3] = [0, 0]\na = [1, 2, 3, 4]\npatch(a)\nprint(a)\n"),
        "[1, 0, 0, 4]"
    );
}

#[test]
fn slice_assign_inside_for_loop_builds_pattern() {
    assert_eq!(
        run_python_one("xs = [0, 0, 0, 0]\nfor i in range(2):\n xs[i:i+1] = [i + 1]\nprint(xs)\n"),
        "[1, 2, 0, 0]"
    );
}

#[test]
fn slice_assign_from_tuple_source() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs[1:2] = (9, 10)\nprint(xs)\n"),
        "[1, 9, 10, 3]"
    );
}

#[test]
fn slice_assign_from_range_object() {
    assert_eq!(
        run_python_one("xs = [0, 0, 0]\nxs[1:2] = list(range(3, 5))\nprint(xs)\n"),
        "[0, 3, 4, 0]"
    );
}

#[test]
fn slice_assign_single_element_slice_same_length() {
    assert_eq!(
        run_python_one("xs = [5, 6, 7]\nxs[1:2] = [60]\nprint(xs)\n"),
        "[5, 60, 7]"
    );
}

#[test]
fn del_single_index_then_slice_assign() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\ndel xs[0]\nxs[1:2] = [30]\nprint(xs)\n"),
        "[2, 30, 4]"
    );
}

#[test]
fn slice_assign_empty_list_noop_on_zero_width_middle() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs[1:1] = []\nprint(xs)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn slice_assign_replaces_all_but_first() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\nxs[1:] = [9]\nprint(xs)\n"),
        "[1, 9]"
    );
}

#[test]
fn slice_assign_replaces_all_but_last() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\nxs[:-1] = [7, 8]\nprint(xs)\n"),
        "[7, 8, 4]"
    );
}

#[test]
fn slice_assign_length_change_updates_len() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs[1:2] = [10, 11, 12]\nprint(len(xs))\n"),
        "5"
    );
}

#[test]
fn slice_assign_after_clear_rebuilds() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs.clear()\nxs[:] = [4, 5]\nprint(xs)\n"),
        "[4, 5]"
    );
}

#[test]
fn slice_assign_copy_from_other_list_segment() {
    assert_eq!(
        run_python_one("a = [1, 2, 3, 4]\nb = [0, 0, 0]\nb[1:3] = a[1:3]\nprint(b)\n"),
        "[0, 2, 3, 0]"
    );
}

#[test]
fn slice_assign_with_list_comprehension_source() {
    assert_eq!(
        run_python_one("xs = [0, 0, 0, 0]\nxs[1:3] = [n * 10 for n in range(2)]\nprint(xs)\n"),
        "[0, 0, 10, 0]"
    );
}

#[test]
fn del_slice_then_append() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3, 4]\ndel xs[1:3]\nxs.append(9)\nprint(xs)\n"),
        "[1, 4, 9]"
    );
}

#[test]
fn slice_assign_on_list_of_lists_row() {
    assert_eq!(
        run_python_one("grid = [[1], [2], [3]]\ngrid[1:2] = [[20, 21]]\nprint(grid)\n"),
        "[[1], [20, 21], [3]]"
    );
}
