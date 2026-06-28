use crate::helpers::{run_print, run_python_one};

#[test]
fn eq_integers_same_value() {
    assert_eq!(run_print("3 == 3"), "true");
}

#[test]
fn eq_integers_different_value() {
    assert_eq!(run_print("3 == 4"), "false");
}

#[test]
fn ne_integers_different_value() {
    assert_eq!(run_print("3 != 4"), "true");
}

#[test]
fn ne_integers_same_value() {
    assert_eq!(run_print("5 != 5"), "false");
}

#[test]
fn lt_smaller_than_larger() {
    assert_eq!(run_print("1 < 2"), "true");
}

#[test]
fn lt_equal_values() {
    assert_eq!(run_print("2 < 2"), "false");
}

#[test]
fn gt_larger_than_smaller() {
    assert_eq!(run_print("5 > 3"), "true");
}

#[test]
fn gt_equal_values() {
    assert_eq!(run_print("4 > 4"), "false");
}

#[test]
fn le_equal_values() {
    assert_eq!(run_print("7 <= 7"), "true");
}

#[test]
fn le_smaller_value() {
    assert_eq!(run_print("2 <= 9"), "true");
}

#[test]
fn le_greater_value() {
    assert_eq!(run_print("9 <= 2"), "false");
}

#[test]
fn ge_equal_values() {
    assert_eq!(run_print("10 >= 10"), "true");
}

#[test]
fn ge_greater_value() {
    assert_eq!(run_print("8 >= 3"), "true");
}

#[test]
fn ge_smaller_value() {
    assert_eq!(run_print("1 >= 5"), "false");
}

#[test]
fn eq_strings_same_content() {
    assert_eq!(run_print("'abc' == 'abc'"), "true");
}

#[test]
fn ne_strings_different_content() {
    assert_eq!(run_print("'abc' != 'xyz'"), "true");
}

#[test]
fn lt_string_lexicographic_order() {
    assert_eq!(run_print("'apple' < 'banana'"), "true");
}

#[test]
fn eq_lists_same_elements() {
    assert_eq!(run_print("[1, 2] == [1, 2]"), "true");
}

#[test]
fn ne_lists_different_elements() {
    assert_eq!(run_print("[1, 2] != [1, 3]"), "true");
}

#[test]
fn in_list_member_present() {
    assert_eq!(run_print("2 in [1, 2, 3]"), "true");
}

#[test]
fn in_list_member_absent() {
    assert_eq!(run_print("9 in [1, 2, 3]"), "false");
}

#[test]
fn not_in_list_member_absent() {
    assert_eq!(run_print("4 not in [1, 2, 3]"), "true");
}

#[test]
fn not_in_list_member_present() {
    assert_eq!(run_print("2 not in [1, 2, 3]"), "false");
}

#[test]
fn in_string_substring_present() {
    assert_eq!(run_print("'ell' in 'hello'"), "true");
}

#[test]
fn not_in_string_substring_absent() {
    assert_eq!(run_print("'z' not in 'hello'"), "true");
}

#[test]
fn in_dict_key_present() {
    assert_eq!(run_print("'a' in {'a': 1, 'b': 2}"), "true");
}

#[test]
fn not_in_dict_key_absent() {
    assert_eq!(run_print("'z' not in {'a': 1}"), "true");
}

#[test]
fn is_same_none_object() {
    assert_eq!(run_python_one("x = None\nprint(x is None)\n"), "true");
}

#[test]
fn is_not_different_objects() {
    assert_eq!(run_python_one("x = []\nprint(x is [])\n"), "false");
}

#[test]
fn is_same_list_reference() {
    assert_eq!(run_python_one("a = [1, 2]\nb = a\nprint(a is b)\n"), "true");
}

#[test]
fn is_not_none_for_integer() {
    assert_eq!(run_python_one("x = 0\nprint(x is not None)\n"), "true");
}

#[test]
fn chained_lt_all_true() {
    assert_eq!(run_print("1 < 2 < 3"), "true");
}

#[test]
fn chained_lt_middle_fails() {
    assert_eq!(run_print("1 < 3 < 2"), "false");
}

#[test]
fn chained_le_ge_range_check_inside() {
    assert_eq!(run_print("0 <= 5 <= 10"), "true");
}

#[test]
fn chained_le_ge_range_check_outside() {
    assert_eq!(run_print("0 <= 15 <= 10"), "false");
}

#[test]
fn chained_eq_all_equal() {
    assert_eq!(run_print("4 == 4 == 4"), "true");
}

#[test]
fn chained_mixed_comparison() {
    assert_eq!(run_print("1 < 2 == 2"), "true");
}

#[test]
fn chained_ne_and_lt() {
    assert_eq!(run_print("3 != 4 < 5"), "true");
}

#[test]
fn compare_floats_equality() {
    assert_eq!(run_print("1.5 == 1.5"), "true");
}

#[test]
fn compare_negative_numbers() {
    assert_eq!(run_print("-3 < -1"), "true");
}
