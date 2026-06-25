use crate::helpers::{run_print, run_python_one};

#[test]
fn enumerate_default_starts_at_zero() {
    assert_eq!(run_print("list(enumerate(['a', 'b']))"), "[(0, 'a'), (1, 'b')]");
}

#[test]
fn enumerate_start_one() {
    assert_eq!(run_print("list(enumerate(['a', 'b'], start=1))"), "[(1, 'a'), (2, 'b')]");
}

#[test]
fn enumerate_start_negative() {
    assert_eq!(run_print("list(enumerate(['x'], start=-1))"), "[(-1, 'x')]");
}

#[test]
fn enumerate_on_range() {
    assert_eq!(run_print("list(enumerate(range(3)))"), "[(0, 0), (1, 1), (2, 2)]");
}

#[test]
fn enumerate_unpack_in_loop() {
    assert_eq!(
        run_python_one("s = ''\nfor i, ch in enumerate('ab'):\n s += str(i) + ch\nprint(s)\n"),
        "0a1b"
    );
}

#[test]
fn enumerate_empty_iterable() {
    assert_eq!(run_print("list(enumerate([]))"), "[]");
}

#[test]
fn enumerate_start_large() {
    assert_eq!(run_print("list(enumerate(['a'], start=100))"), "[(100, 'a')]");
}

#[test]
fn enumerate_on_tuple() {
    assert_eq!(run_print("list(enumerate((10, 20)))"), "[(0, 10), (1, 20)]");
}

#[test]
fn enumerate_on_string_chars() {
    assert_eq!(run_print("list(enumerate('hi'))"), "[(0, 'h'), (1, 'i')]");
}

#[test]
fn enumerate_start_zero_explicit() {
    assert_eq!(run_print("list(enumerate(['z'], start=0))"), "[(0, 'z')]");
}

#[test]
fn enumerate_index_used_in_expression() {
    assert_eq!(
        run_print("[i * 10 + v for i, v in enumerate([1, 2, 3])]"),
        "[1, 12, 23]"
    );
}

#[test]
fn enumerate_with_list_comprehension_filter() {
    assert_eq!(
        run_print("[i for i, ch in enumerate('aba') if ch == 'a']"),
        "[0, 2]"
    );
}

#[test]
fn enumerate_dict_keys_by_index() {
    assert_eq!(
        run_python_one("d = {'b': 2, 'a': 1}\nprint(list(enumerate(d))[0][1])\n"),
        "b"
    );
}

#[test]
fn enumerate_start_two_step_values() {
    assert_eq!(
        run_print("list(enumerate([5, 6, 7], start=2))"),
        "[(2, 5), (3, 6), (4, 7)]"
    );
}

#[test]
fn enumerate_materialized_len() {
    assert_eq!(run_print("len(list(enumerate('abcd')))"), "4");
}

#[test]
fn enumerate_first_index_only() {
    assert_eq!(
        run_python_one("for i, _ in enumerate(['only']):\n print(i)\n"),
        "0"
    );
}

#[test]
fn enumerate_on_list_of_lists() {
    assert_eq!(
        run_print("list(enumerate([[1], [2, 3]]))"),
        "[(0, [1]), (1, [2, 3])]"
    );
}

#[test]
fn enumerate_start_with_single_element() {
    assert_eq!(run_print("list(enumerate([99], start=5))"), "[(5, 99)]");
}

#[test]
fn enumerate_zip_parallel() {
    assert_eq!(
        run_print("list(zip(enumerate(['a', 'b']), ['x', 'y']))"),
        "[((0, 'a'), 'x'), ((1, 'b'), 'y')]"
    );
}

#[test]
fn enumerate_bool_values() {
    assert_eq!(
        run_print("list(enumerate([True, False]))"),
        "[(0, True), (1, False)]"
    );
}

#[test]
fn enumerate_start_at_ten_count_three() {
    assert_eq!(
        run_print("[i for i, _ in enumerate(range(3), start=10)]"),
        "[10, 11, 12]"
    );
}

#[test]
fn enumerate_for_else_not_triggered() {
    assert_eq!(
        run_python_one("for i, v in enumerate([1]):\n pass\nelse:\n print(i)\n"),
        "0"
    );
}

#[test]
fn enumerate_break_keeps_partial() {
    assert_eq!(
        run_python_one("out = []\nfor i, v in enumerate([1, 2, 3]):\n if i == 1:\n  break\n out.append(i)\nprint(out)\n"),
        "[0]"
    );
}

#[test]
fn enumerate_continue_skips_index() {
    assert_eq!(
        run_python_one("out = []\nfor i, v in enumerate([1, 2, 3]):\n if i == 1:\n  continue\n out.append(i)\nprint(out)\n"),
        "[0, 2]"
    );
}

#[test]
fn enumerate_on_bytes_like_int_list() {
    assert_eq!(
        run_print("list(enumerate([65, 66, 67]))"),
        "[(0, 65), (1, 66), (2, 67)]"
    );
}

#[test]
fn enumerate_nested_loop_outer_index() {
    assert_eq!(
        run_python_one("out = []\nfor i, row in enumerate([[1], [2, 3]]):\n out.append(i)\nprint(out)\n"),
        "[0, 1]"
    );
}

#[test]
fn enumerate_start_negative_with_two_items() {
    assert_eq!(
        run_print("list(enumerate(['a', 'b'], start=-2))"),
        "[(-2, 'a'), (-1, 'b')]"
    );
}

#[test]
fn enumerate_value_not_used() {
    assert_eq!(
        run_print("len([i for i, _ in enumerate('hello')])"),
        "5"
    );
}

#[test]
fn enumerate_with_start_in_fstring() {
    assert_eq!(
        run_python_one("pairs = list(enumerate(['x'], start=7))\nprint(f'{pairs[0][0]}')\n"),
        "7"
    );
}

#[test]
fn enumerate_generator_expr_sum_indices() {
    assert_eq!(
        run_print("sum(i for i, _ in enumerate('abc'))"),
        "3"
    );
}
