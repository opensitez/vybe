use crate::helpers::{run_print, run_python_one};

#[test]
fn none_singleton_print() {
    assert_eq!(run_print("None"), "None");
}

#[test]
fn none_is_singleton() {
    assert_eq!(run_print("None is None"), "True");
}

#[test]
fn none_equality() {
    assert_eq!(run_print("None == None"), "True");
}

#[test]
fn none_bool_false() {
    assert_eq!(run_print("bool(None)"), "False");
}

#[test]
fn none_not_none_true() {
    assert_eq!(run_print("not None"), "True");
}

#[test]
fn none_is_not_zero() {
    assert_eq!(run_print("None is 0"), "False");
}

#[test]
fn none_is_not_empty_string() {
    assert_eq!(run_print("None is ''"), "False");
}

#[test]
fn none_is_not_empty_list() {
    assert_eq!(run_print("None is []"), "False");
}

#[test]
fn none_equals_only_none() {
    assert_eq!(run_print("None == 0"), "False");
}

#[test]
fn none_type_name() {
    assert_eq!(run_print("type(None).__name__"), "NoneType");
}

#[test]
fn truthy_nonempty_string() {
    assert_eq!(run_print("bool('a')"), "True");
}

#[test]
fn falsy_empty_string() {
    assert_eq!(run_print("bool('')"), "False");
}

#[test]
fn truthy_zero_int_is_falsy() {
    assert_eq!(run_print("bool(0)"), "False");
}

#[test]
fn truthy_positive_int() {
    assert_eq!(run_print("bool(1)"), "True");
}

#[test]
fn truthy_negative_int() {
    assert_eq!(run_print("bool(-1)"), "True");
}

#[test]
fn falsy_zero_float() {
    assert_eq!(run_print("bool(0.0)"), "False");
}

#[test]
fn truthy_float() {
    assert_eq!(run_print("bool(0.1)"), "True");
}

#[test]
fn falsy_empty_list() {
    assert_eq!(run_print("bool([])"), "False");
}

#[test]
fn truthy_nonempty_list() {
    assert_eq!(run_print("bool([0])"), "True");
}

#[test]
fn falsy_empty_tuple() {
    assert_eq!(run_print("bool(())"), "False");
}

#[test]
fn truthy_nonempty_tuple() {
    assert_eq!(run_print("bool((0,))"), "True");
}

#[test]
fn falsy_empty_dict() {
    assert_eq!(run_print("bool({})"), "False");
}

#[test]
fn truthy_nonempty_dict() {
    assert_eq!(run_print("bool({0: 0})"), "True");
}

#[test]
fn falsy_empty_set() {
    assert_eq!(run_print("bool(set())"), "False");
}

#[test]
fn truthy_nonempty_set() {
    assert_eq!(run_print("bool({0})"), "True");
}

#[test]
fn truthy_true_singleton() {
    assert_eq!(run_print("bool(True)"), "True");
}

#[test]
fn falsy_false_singleton() {
    assert_eq!(run_print("bool(False)"), "False");
}

#[test]
fn and_short_circuit_falsy() {
    assert_eq!(run_print("0 and 99"), "0");
}

#[test]
fn and_short_circuit_truthy_returns_last() {
    assert_eq!(run_print("1 and 2 and 3"), "3");
}

#[test]
fn or_short_circuit_truthy() {
    assert_eq!(run_print("1 or 99"), "1");
}

#[test]
fn or_short_circuit_falsy_returns_last() {
    assert_eq!(run_print("0 or '' or 7"), "7");
}

#[test]
fn not_truthy_to_false() {
    assert_eq!(run_print("not 1"), "False");
}

#[test]
fn not_falsy_to_true() {
    assert_eq!(run_print("not 0"), "True");
}

#[test]
fn if_truthy_branch() {
    assert_eq!(
        run_python_one("x = 1\nprint('yes' if x else 'no')\n"),
        "yes"
    );
}

#[test]
fn if_falsy_branch() {
    assert_eq!(run_python_one("x = 0\nprint('yes' if x else 'no')\n"), "no");
}

#[test]
fn filter_none_removes_falsy() {
    assert_eq!(
        run_print("list(filter(None, [0, 1, '', 'a', None]))"),
        "[1, 'a']"
    );
}

#[test]
fn all_empty_true() {
    assert_eq!(run_print("all([])"), "True");
}

#[test]
fn any_empty_false() {
    assert_eq!(run_print("any([])"), "False");
}

#[test]
fn all_with_zero_false() {
    assert_eq!(run_print("all([1, 2, 0])"), "False");
}

#[test]
fn any_with_truthy_true() {
    assert_eq!(run_print("any([0, 0, 1])"), "True");
}

#[test]
fn none_return_default() {
    assert_eq!(run_python_one("def f():\n pass\nprint(f())\n"), "None");
}

#[test]
fn compare_none_with_is() {
    assert_eq!(run_python_one("x = None\nprint(x is None)\n"), "True");
}

#[test]
fn compare_none_equality_safe() {
    assert_eq!(run_python_one("x = None\nprint(x == None)\n"), "True");
}

#[test]
fn truthy_custom_nonzero_len() {
    assert_eq!(
        run_python_one("class C:\n def __len__(self):\n  return 1\nprint(bool(C()))\n"),
        "True"
    );
}

#[test]
fn falsy_custom_zero_len() {
    assert_eq!(
        run_python_one("class C:\n def __len__(self):\n  return 0\nprint(bool(C()))\n"),
        "False"
    );
}

#[test]
fn truthy_custom_bool_dunder() {
    assert_eq!(
        run_python_one("class C:\n def __bool__(self):\n  return True\nprint(bool(C()))\n"),
        "True"
    );
}
