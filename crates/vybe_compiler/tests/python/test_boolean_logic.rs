use crate::helpers::{run_print, run_python, run_python_one};

#[test]
fn and_both_true_returns_true() {
    assert_eq!(run_print("True and True"), "True");
}

#[test]
fn and_left_false_returns_false() {
    assert_eq!(run_print("False and True"), "False");
}

#[test]
fn and_right_false_returns_false() {
    assert_eq!(run_print("True and False"), "False");
}

#[test]
fn and_both_false_returns_false() {
    assert_eq!(run_print("False and False"), "False");
}

#[test]
fn or_both_true_returns_true() {
    assert_eq!(run_print("True or True"), "True");
}

#[test]
fn or_left_true_returns_true() {
    assert_eq!(run_print("True or False"), "True");
}

#[test]
fn or_right_true_returns_true() {
    assert_eq!(run_print("False or True"), "True");
}

#[test]
fn or_both_false_returns_false() {
    assert_eq!(run_print("False or False"), "False");
}

#[test]
fn not_true_returns_false() {
    assert_eq!(run_print("not True"), "False");
}

#[test]
fn not_false_returns_true() {
    assert_eq!(run_print("not False"), "True");
}

#[test]
fn and_short_circuit_skips_right_operand() {
    assert_eq!(
        run_python("def boom():\n    print('boom')\n    return True\nprint(False and boom())\n",),
        vec!["False"]
    );
}

#[test]
fn or_short_circuit_skips_right_operand() {
    assert_eq!(
        run_python("def boom():\n    print('boom')\n    return True\nprint(True or boom())\n",),
        vec!["True"]
    );
}

#[test]
fn and_short_circuit_returns_left_when_falsy() {
    assert_eq!(run_print("0 and 99"), "0");
}

#[test]
fn and_short_circuit_returns_right_when_left_truthy() {
    assert_eq!(run_print("1 and 99"), "99");
}

#[test]
fn or_short_circuit_returns_left_when_truthy() {
    assert_eq!(run_print("42 or 0"), "42");
}

#[test]
fn or_short_circuit_returns_right_when_left_falsy() {
    assert_eq!(run_print("0 or 7"), "7");
}

#[test]
fn chained_and_requires_all_truthy() {
    assert_eq!(run_print("1 and 2 and 3"), "3");
}

#[test]
fn chained_and_stops_at_first_falsy() {
    assert_eq!(run_print("1 and 0 and 3"), "0");
}

#[test]
fn chained_or_returns_first_truthy() {
    assert_eq!(run_print("0 or '' or 'ok'"), "ok");
}

#[test]
fn not_inverts_truthy_value() {
    assert_eq!(run_print("not 1"), "False");
}

#[test]
fn not_inverts_falsy_zero() {
    assert_eq!(run_print("not 0"), "True");
}

#[test]
fn bool_true_literal() {
    assert_eq!(run_print("bool(True)"), "True");
}

#[test]
fn bool_false_literal() {
    assert_eq!(run_print("bool(False)"), "False");
}

#[test]
fn bool_zero_is_false() {
    assert_eq!(run_print("bool(0)"), "False");
}

#[test]
fn bool_positive_int_is_true() {
    assert_eq!(run_print("bool(42)"), "True");
}

#[test]
fn bool_negative_int_is_true() {
    assert_eq!(run_print("bool(-1)"), "True");
}

#[test]
fn bool_empty_string_is_false() {
    assert_eq!(run_print("bool('')"), "False");
}

#[test]
fn bool_nonempty_string_is_true() {
    assert_eq!(run_print("bool('x')"), "True");
}

#[test]
fn bool_empty_list_is_false() {
    assert_eq!(run_print("bool([])"), "False");
}

#[test]
fn bool_nonempty_list_is_true() {
    assert_eq!(run_print("bool([0])"), "True");
}

#[test]
fn bool_empty_tuple_is_false() {
    assert_eq!(run_print("bool(())"), "False");
}

#[test]
fn bool_nonempty_tuple_is_true() {
    assert_eq!(run_print("bool((1,))"), "True");
}

#[test]
fn bool_empty_dict_is_false() {
    assert_eq!(run_print("bool({})"), "False");
}

#[test]
fn bool_nonempty_dict_is_true() {
    assert_eq!(run_print("bool({'a': 1})"), "True");
}

#[test]
fn bool_none_is_false() {
    assert_eq!(run_print("bool(None)"), "False");
}

#[test]
fn if_empty_list_branch_is_false() {
    assert_eq!(
        run_python_one("if []:\n    print('yes')\nelse:\n    print('no')\n"),
        "no"
    );
}

#[test]
fn if_nonempty_list_branch_is_true() {
    assert_eq!(
        run_python_one("if [1]:\n    print('yes')\nelse:\n    print('no')\n"),
        "yes"
    );
}

#[test]
fn if_zero_branch_is_false() {
    assert_eq!(
        run_python_one("if 0:\n    print('yes')\nelse:\n    print('no')\n"),
        "no"
    );
}

#[test]
fn if_nonempty_string_branch_is_true() {
    assert_eq!(
        run_python_one("if 'hi':\n    print('yes')\nelse:\n    print('no')\n"),
        "yes"
    );
}

#[test]
fn double_not_restores_truthiness() {
    assert_eq!(run_print("not not 5"), "True");
}
