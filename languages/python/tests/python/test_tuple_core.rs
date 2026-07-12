use crate::helpers::{run_print, run_python_one};

#[test]
fn tuple_literal_two_elements() {
    assert_eq!(run_print("(1, 2)"), "(1, 2)");
}

#[test]
fn tuple_single_element_trailing_comma() {
    assert_eq!(run_print("(1,)"), "(1,)");
}

#[test]
fn tuple_empty() {
    assert_eq!(run_print("()"), "()");
}

#[test]
fn tuple_index_first() {
    assert_eq!(run_print("(10, 20, 30)[0]"), "10");
}

#[test]
fn tuple_index_last() {
    assert_eq!(run_print("(10, 20, 30)[-1]"), "30");
}

#[test]
fn tuple_slice_middle() {
    assert_eq!(run_print("(0, 1, 2, 3)[1:3]"), "(1, 2)");
}

#[test]
fn tuple_concat() {
    assert_eq!(run_print("(1, 2) + (3,)"), "(1, 2, 3)");
}

#[test]
fn tuple_repeat() {
    assert_eq!(run_print("(1,) * 4"), "(1, 1, 1, 1)");
}

#[test]
fn tuple_unpack_two() {
    assert_eq!(run_python_one("a, b = (1, 2)\nprint(a, b)\n"), "1 2");
}

#[test]
fn tuple_unpack_three() {
    assert_eq!(run_python_one("a, b, c = (1, 2, 3)\nprint(c)\n"), "3");
}

#[test]
fn tuple_nested() {
    assert_eq!(run_print("((1, 2), (3, 4))"), "((1, 2), (3, 4))");
}

#[test]
fn tuple_in_list() {
    assert_eq!(run_print("[(1, 2), (3, 4)]"), "[(1, 2), (3, 4)]");
}

#[test]
fn tuple_as_dict_key() {
    assert_eq!(run_print("{(1, 2): 'pair'}[(1, 2)]"), "pair");
}

#[test]
fn tuple_equality() {
    assert_eq!(run_print("(1, 2) == (1, 2)"), "True");
}

#[test]
fn tuple_inequality() {
    assert_eq!(run_print("(1, 2) == (2, 1)"), "False");
}

#[test]
fn tuple_less_than_lexicographic() {
    assert_eq!(run_print("(1, 2) < (1, 3)"), "True");
}

#[test]
fn tuple_length() {
    assert_eq!(run_print("len((1, 2, 3))"), "3");
}

#[test]
fn tuple_contains() {
    assert_eq!(run_print("2 in (1, 2, 3)"), "True");
}

#[test]
fn tuple_count_method() {
    assert_eq!(run_print("(1, 2, 1, 1).count(1)"), "3");
}

#[test]
fn tuple_index_method() {
    assert_eq!(run_print("(1, 2, 3).index(2)"), "1");
}

#[test]
fn tuple_from_list() {
    assert_eq!(run_print("tuple([1, 2])"), "(1, 2)");
}

#[test]
fn tuple_from_string() {
    assert_eq!(run_print("tuple('ab')"), "('a', 'b')");
}

#[test]
fn tuple_iter_sum() {
    assert_eq!(run_print("sum((1, 2, 3))"), "6");
}

#[test]
fn tuple_unpack_star_middle() {
    assert_eq!(
        run_python_one("first, *mid, last = (1, 2, 3, 4)\nprint(first, mid, last)\n"),
        "1 [2, 3] 4"
    );
}

#[test]
fn tuple_unpack_star_start() {
    assert_eq!(
        run_python_one("a, *rest = (1, 2, 3)\nprint(a, rest)\n"),
        "1 [2, 3]"
    );
}

#[test]
fn tuple_compare_unequal_length() {
    assert_eq!(run_print("(1,) < (1, 2)"), "True");
}

#[test]
fn tuple_bool_nonempty_true() {
    assert_eq!(run_print("bool((0,))"), "True");
}

#[test]
fn tuple_bool_empty_false() {
    assert_eq!(run_print("bool(())"), "False");
}

#[test]
fn tuple_mixed_types() {
    assert_eq!(run_print("(1, 'a', None)"), "(1, 'a', None)");
}

#[test]
fn tuple_hashable_in_set() {
    assert_eq!(run_print("len({(1, 2), (1, 2)})"), "1");
}

#[test]
fn tuple_return_from_function() {
    assert_eq!(
        run_python_one("def pair():\n return 1, 2\nprint(pair())\n"),
        "(1, 2)"
    );
}

#[test]
fn tuple_swap_via_unpack() {
    assert_eq!(
        run_python_one("a, b = 1, 2\na, b = b, a\nprint(a, b)\n"),
        "2 1"
    );
}

#[test]
fn tuple_zip_result() {
    assert_eq!(
        run_print("list(zip([1, 2], ['a', 'b']))"),
        "[(1, 'a'), (2, 'b')]"
    );
}

#[test]
fn tuple_enumerate_result() {
    assert_eq!(run_print("list(enumerate(['x']))[0]"), "(0, 'x')");
}

#[test]
fn tuple_min_max() {
    assert_eq!(run_print("[min((3, 1, 2)), max((3, 1, 2))]"), "[1, 3]");
}

#[test]
fn tuple_any_all() {
    assert_eq!(run_print("[any((0, 1)), all((1, 2))]"), "[True, True]");
}

#[test]
fn tuple_slice_step() {
    assert_eq!(run_print("(0, 1, 2, 3, 4)[::2]"), "(0, 2, 4)");
}

#[test]
fn tuple_reverse_slice() {
    assert_eq!(run_print("(1, 2, 3)[::-1]"), "(3, 2, 1)");
}

#[test]
fn tuple_compare_nested() {
    assert_eq!(run_print("((1, 2),) == ((1, 2),)"), "True");
}

#[test]
fn tuple_in_operator() {
    assert_eq!(run_print("(1, 2) in [(1, 2), (3, 4)]"), "True");
}

#[test]
fn tuple_constructor_no_args() {
    assert_eq!(run_print("tuple()"), "()");
}

#[test]
fn tuple_constructor_from_range() {
    assert_eq!(run_print("tuple(range(3))"), "(0, 1, 2)");
}

#[test]
fn tuple_unpack_in_for_loop() {
    assert_eq!(
        run_python_one("total = 0\nfor a, b in [(1, 2), (3, 4)]:\n total += a + b\nprint(total)\n"),
        "10"
    );
}

#[test]
fn tuple_equality_with_list_false() {
    assert_eq!(run_print("(1, 2) == [1, 2]"), "False");
}

#[test]
fn tuple_identity_not_same_object() {
    assert_eq!(run_print("(1,) is (1,)"), "False");
}
