use crate::helpers::{run_print, run_python_one};

#[test]
fn sorted_ascending_integers() {
    assert_eq!(run_print("sorted([3, 1, 2])"), "[1, 2, 3]");
}

#[test]
fn sorted_descending_with_reverse_true() {
    assert_eq!(run_print("sorted([3, 1, 2], reverse=True)"), "[3, 2, 1]");
}

#[test]
fn sorted_empty_list_returns_empty() {
    assert_eq!(run_print("sorted([])"), "[]");
}

#[test]
fn sorted_single_element_unchanged() {
    assert_eq!(run_print("sorted([42])"), "[42]");
}

#[test]
fn sorted_strings_alphabetical() {
    assert_eq!(
        run_print("sorted(['banana', 'apple', 'cherry'])"),
        "['apple', 'banana', 'cherry']"
    );
}

#[test]
fn sorted_does_not_mutate_original() {
    assert_eq!(
        run_python_one("xs = [3, 1, 2]\nsorted(xs)\nprint(xs)\n"),
        "[3, 1, 2]"
    );
}

#[test]
fn sorted_negative_numbers() {
    assert_eq!(run_print("sorted([-1, -3, 0, 2])"), "[-3, -1, 0, 2]");
}

#[test]
fn sorted_duplicate_values_stable_order() {
    assert_eq!(run_print("sorted([2, 1, 2, 1])"), "[1, 1, 2, 2]");
}

#[test]
fn sorted_mixed_sign_floats() {
    assert_eq!(run_print("sorted([1.5, -0.5, 0.0])"), "[-0.5, 0.0, 1.5]");
}

#[test]
fn sorted_tuple_input_returns_list() {
    assert_eq!(run_print("sorted((3, 1, 2))"), "[1, 2, 3]");
}

#[test]
fn sorted_range_object_via_list() {
    assert_eq!(
        run_print("sorted(list(range(5, 0, -1)))"),
        "[1, 2, 3, 4, 5]"
    );
}

#[test]
fn min_of_integer_list() {
    assert_eq!(run_print("min([5, 2, 8, 1])"), "1");
}

#[test]
fn max_of_integer_list() {
    assert_eq!(run_print("max([5, 2, 8, 1])"), "8");
}

#[test]
fn min_of_two_arguments() {
    assert_eq!(run_print("min(10, 3)"), "3");
}

#[test]
fn max_of_two_arguments() {
    assert_eq!(run_print("max(10, 3)"), "10");
}

#[test]
fn min_empty_raises_or_handles() {
    assert_eq!(
        run_python_one("try:\n print(min([]))\nexcept Exception as e:\n print(type(e).__name__)\n"),
        "ValueError"
    );
}

#[test]
fn max_empty_raises_value_error() {
    assert_eq!(
        run_python_one("try:\n print(max([]))\nexcept Exception as e:\n print(type(e).__name__)\n"),
        "ValueError"
    );
}

#[test]
fn min_string_sequence() {
    assert_eq!(run_print("min(['dog', 'ant', 'cat'])"), "ant");
}

#[test]
fn max_string_sequence() {
    assert_eq!(run_print("max(['dog', 'ant', 'cat'])"), "dog");
}

#[test]
fn min_with_default_on_empty() {
    assert_eq!(run_print("min([], default=99)"), "99");
}

#[test]
fn max_with_default_on_empty() {
    assert_eq!(run_print("max([], default=-1)"), "-1");
}

#[test]
fn sum_empty_list_is_zero() {
    assert_eq!(run_print("sum([])"), "0");
}

#[test]
fn sum_single_element() {
    assert_eq!(run_print("sum([7])"), "7");
}

#[test]
fn sum_positive_integers() {
    assert_eq!(run_print("sum([1, 2, 3, 4])"), "10");
}

#[test]
fn sum_with_start_value() {
    assert_eq!(run_print("sum([1, 2, 3], 10)"), "16");
}

#[test]
fn sum_negative_numbers() {
    assert_eq!(run_print("sum([-1, -2, -3])"), "-6");
}

#[test]
fn sum_mixed_positive_negative() {
    assert_eq!(run_print("sum([5, -2, 3, -1])"), "5");
}

#[test]
fn reversed_list_iterator_materialized() {
    assert_eq!(run_print("list(reversed([1, 2, 3]))"), "[3, 2, 1]");
}

#[test]
fn reversed_empty_list() {
    assert_eq!(run_print("list(reversed([]))"), "[]");
}

#[test]
fn reversed_single_item() {
    assert_eq!(run_print("list(reversed([99]))"), "[99]");
}

#[test]
fn reversed_string_chars() {
    assert_eq!(run_print("list(reversed('abc'))"), "['c', 'b', 'a']");
}

#[test]
fn reversed_does_not_mutate_list() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nlist(reversed(xs))\nprint(xs)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn list_sort_in_place_ascending() {
    assert_eq!(
        run_python_one("xs = [3, 1, 2]\nxs.sort()\nprint(xs)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn list_sort_reverse_true() {
    assert_eq!(
        run_python_one("xs = [3, 1, 2]\nxs.sort(reverse=True)\nprint(xs)\n"),
        "[3, 2, 1]"
    );
}

#[test]
fn list_sort_returns_none() {
    assert_eq!(
        run_python_one("xs = [2, 1]\nr = xs.sort()\nprint(r)\n"),
        "None"
    );
}

#[test]
fn sorted_key_abs_on_negatives() {
    assert_eq!(run_print("sorted([-3, 1, -2], key=abs)"), "[1, -2, -3]");
}

#[test]
fn sorted_key_len_for_strings() {
    assert_eq!(
        run_print("sorted(['aaa', 'b', 'cc'], key=len)"),
        "['b', 'cc', 'aaa']"
    );
}

#[test]
fn min_of_tuple_literal() {
    assert_eq!(run_print("min((9, 4, 7))"), "4");
}

#[test]
fn max_of_tuple_literal() {
    assert_eq!(run_print("max((9, 4, 7))"), "9");
}

#[test]
fn sum_of_range_list() {
    assert_eq!(run_print("sum(list(range(1, 6)))"), "15");
}

#[test]
fn sorted_bool_values_false_before_true() {
    assert_eq!(
        run_print("sorted([True, False, True])"),
        "[False, False, True]"
    );
}

#[test]
fn min_bool_false_is_less_than_true() {
    assert_eq!(run_print("min([True, False])"), "False");
}

#[test]
fn max_bool_true_beats_false() {
    assert_eq!(run_print("max([True, False])"), "True");
}

#[test]
fn sorted_nested_lists_by_first_element() {
    assert_eq!(
        run_print("sorted([[2, 0], [1, 9], [3, 3]], key=lambda x: x[0])"),
        "[[1, 9], [2, 0], [3, 3]]"
    );
}

#[test]
fn reversed_range_via_list() {
    assert_eq!(run_print("list(reversed(range(4)))"), "[3, 2, 1, 0]");
}

#[test]
fn sum_floats_with_fractions() {
    assert_eq!(run_print("sum([0.1, 0.2, 0.3])"), "0.6");
}

#[test]
fn sorted_three_equal_elements() {
    assert_eq!(run_print("sorted([5, 5, 5])"), "[5, 5, 5]");
}

#[test]
fn min_varargs_many_integers() {
    assert_eq!(run_print("min(8, 2, 5, 1, 9)"), "1");
}

#[test]
fn max_varargs_many_integers() {
    assert_eq!(run_print("max(8, 2, 5, 1, 9)"), "9");
}

#[test]
fn sorted_mixed_int_float_coercion() {
    assert_eq!(run_print("sorted([2, 1.5, 3])"), "[1.5, 2, 3]");
}
