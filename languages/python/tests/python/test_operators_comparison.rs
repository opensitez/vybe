use crate::helpers::{run_print, run_python_one};

#[test]
fn eq_integers_same_value() {
    assert_eq!(run_print("3 == 3"), "True");
}

#[test]
fn eq_integers_different_value() {
    assert_eq!(run_print("3 == 4"), "False");
}

#[test]
fn ne_integers_different_value() {
    assert_eq!(run_print("3 != 4"), "True");
}

#[test]
fn ne_integers_same_value() {
    assert_eq!(run_print("5 != 5"), "False");
}

#[test]
fn lt_smaller_than_larger() {
    assert_eq!(run_print("1 < 2"), "True");
}

#[test]
fn lt_equal_values() {
    assert_eq!(run_print("2 < 2"), "False");
}

#[test]
fn gt_larger_than_smaller() {
    assert_eq!(run_print("5 > 3"), "True");
}

#[test]
fn gt_equal_values() {
    assert_eq!(run_print("4 > 4"), "False");
}

#[test]
fn le_equal_values() {
    assert_eq!(run_print("7 <= 7"), "True");
}

#[test]
fn le_smaller_value() {
    assert_eq!(run_print("2 <= 9"), "True");
}

#[test]
fn le_greater_value() {
    assert_eq!(run_print("9 <= 2"), "False");
}

#[test]
fn ge_equal_values() {
    assert_eq!(run_print("10 >= 10"), "True");
}

#[test]
fn ge_greater_value() {
    assert_eq!(run_print("8 >= 3"), "True");
}

#[test]
fn ge_smaller_value() {
    assert_eq!(run_print("1 >= 5"), "False");
}

#[test]
fn eq_strings_same_content() {
    assert_eq!(run_print("'abc' == 'abc'"), "True");
}

#[test]
fn ne_strings_different_content() {
    assert_eq!(run_print("'abc' != 'xyz'"), "True");
}

#[test]
fn lt_string_lexicographic_order() {
    assert_eq!(run_print("'apple' < 'banana'"), "True");
}

#[test]
fn eq_lists_same_elements() {
    assert_eq!(run_print("[1, 2] == [1, 2]"), "True");
}

#[test]
fn ne_lists_different_elements() {
    assert_eq!(run_print("[1, 2] != [1, 3]"), "True");
}

#[test]
fn in_list_member_present() {
    assert_eq!(run_print("2 in [1, 2, 3]"), "True");
}

#[test]
fn in_list_member_absent() {
    assert_eq!(run_print("9 in [1, 2, 3]"), "False");
}

#[test]
fn not_in_list_member_absent() {
    assert_eq!(run_print("4 not in [1, 2, 3]"), "True");
}

#[test]
fn not_in_list_member_present() {
    assert_eq!(run_print("2 not in [1, 2, 3]"), "False");
}

#[test]
fn in_string_substring_present() {
    assert_eq!(run_print("'ell' in 'hello'"), "True");
}

#[test]
fn not_in_string_substring_absent() {
    assert_eq!(run_print("'z' not in 'hello'"), "True");
}

#[test]
fn in_dict_key_present() {
    assert_eq!(run_print("'a' in {'a': 1, 'b': 2}"), "True");
}

#[test]
fn not_in_dict_key_absent() {
    assert_eq!(run_print("'z' not in {'a': 1}"), "True");
}

#[test]
fn is_same_none_object() {
    assert_eq!(run_python_one("x = None\nprint(x is None)\n"), "True");
}

#[test]
fn is_not_different_objects() {
    assert_eq!(run_python_one("x = []\nprint(x is [])\n"), "False");
}

#[test]
fn is_same_list_reference() {
    assert_eq!(run_python_one("a = [1, 2]\nb = a\nprint(a is b)\n"), "True");
}

#[test]
fn is_not_none_for_integer() {
    assert_eq!(run_python_one("x = 0\nprint(x is not None)\n"), "True");
}

#[test]
fn chained_lt_all_true() {
    assert_eq!(run_print("1 < 2 < 3"), "True");
}

#[test]
fn chained_lt_middle_fails() {
    assert_eq!(run_print("1 < 3 < 2"), "False");
}

#[test]
fn chained_le_ge_range_check_inside() {
    assert_eq!(run_print("0 <= 5 <= 10"), "True");
}

#[test]
fn chained_le_ge_range_check_outside() {
    assert_eq!(run_print("0 <= 15 <= 10"), "False");
}

#[test]
fn chained_eq_all_equal() {
    assert_eq!(run_print("4 == 4 == 4"), "True");
}

#[test]
fn chained_mixed_comparison() {
    assert_eq!(run_print("1 < 2 == 2"), "True");
}

#[test]
fn chained_ne_and_lt() {
    assert_eq!(run_print("3 != 4 < 5"), "True");
}

#[test]
fn compare_floats_equality() {
    assert_eq!(run_print("1.5 == 1.5"), "True");
}

#[test]
fn compare_negative_numbers() {
    assert_eq!(run_print("-3 < -1"), "True");
}

// ── `in` / `not in` across the RUNTIME probe's legs ────────────────────────
//
// Python cannot resolve a receiver's type at compile time in idiomatic code, so
// `x in y` tests `y` at run time and dispatches to that type's `Contains`
// binding — `[builtin_slots.string|array|map] contains`
// (builtinslotplan.md §3i). The list and dict legs were covered above; these
// pin the string leg and the two receivers that fall through to the third.
//
// The map leg is `ecma:object.hasIn`, NOT the platform default own-only
// `dict.has`, because `in` on a Python object must see inherited members. That
// distinction is the reason the probe requires all three to be declared rather
// than defaulting the gaps.

#[test]
fn membership_substring_in_string() {
    assert_eq!(run_print("'wor' in 'hello world'"), "True");
}

#[test]
fn membership_absent_substring_in_string() {
    assert_eq!(run_print("'xyz' in 'hello world'"), "False");
}

#[test]
fn membership_not_in_string() {
    assert_eq!(run_print("'z' not in 'abc'"), "True");
}

#[test]
fn membership_element_in_set() {
    assert_eq!(run_print("2 in {1, 2, 3}"), "True");
}

#[test]
fn membership_absent_element_in_set() {
    assert_eq!(run_print("9 in {1, 2, 3}"), "False");
}

#[test]
fn membership_element_in_tuple() {
    assert_eq!(run_print("5 in (4, 5, 6)"), "True");
}

#[test]
fn membership_not_in_tuple() {
    assert_eq!(run_print("9 not in (4, 5, 6)"), "True");
}
