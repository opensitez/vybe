use crate::helpers::{run_print, run_python_one};

#[test]
fn list_iadd_extends_in_place() {
    assert_eq!(
        run_python_one("xs = [1]\nxs += [2, 3]\nprint(xs)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn list_iadd_empty_no_change() {
    assert_eq!(
        run_python_one("xs = [1, 2]\nxs += []\nprint(xs)\n"),
        "[1, 2]"
    );
}

#[test]
fn list_imul_repeats() {
    assert_eq!(
        run_python_one("xs = [1, 2]\nxs *= 2\nprint(xs)\n"),
        "[1, 2, 1, 2]"
    );
}

#[test]
fn list_imul_zero_clears() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs *= 0\nprint(xs)\n"),
        "[]"
    );
}

#[test]
fn str_iadd_concatenates() {
    assert_eq!(
        run_python_one("s = 'a'\ns += 'bc'\nprint(s)\n"),
        "abc"
    );
}

#[test]
fn str_imul_repeats() {
    assert_eq!(
        run_python_one("s = 'ab'\ns *= 3\nprint(s)\n"),
        "ababab"
    );
}

#[test]
fn str_imul_zero_empty() {
    assert_eq!(
        run_python_one("s = 'hello'\ns *= 0\nprint(repr(s))\n"),
        "''"
    );
}

#[test]
fn int_iadd_increment() {
    assert_eq!(
        run_python_one("n = 10\nn += 5\nprint(n)\n"),
        "15"
    );
}

#[test]
fn int_isub_decrement() {
    assert_eq!(
        run_python_one("n = 10\nn -= 3\nprint(n)\n"),
        "7"
    );
}

#[test]
fn int_imul_scale() {
    assert_eq!(
        run_python_one("n = 4\nn *= 3\nprint(n)\n"),
        "12"
    );
}

#[test]
fn int_ifloordiv_assign() {
    assert_eq!(
        run_python_one("n = 17\nn //= 5\nprint(n)\n"),
        "3"
    );
}

#[test]
fn int_imod_assign() {
    assert_eq!(
        run_python_one("n = 17\nn %= 5\nprint(n)\n"),
        "2"
    );
}

#[test]
fn int_ipow_assign() {
    assert_eq!(
        run_python_one("n = 2\nn **= 4\nprint(n)\n"),
        "16"
    );
}

#[test]
fn float_iadd() {
    assert_eq!(
        run_python_one("x = 1.5\nx += 2.5\nprint(x)\n"),
        "4.0"
    );
}

#[test]
fn float_idiv_assign() {
    assert_eq!(
        run_python_one("x = 10.0\nx /= 4\nprint(x)\n"),
        "2.5"
    );
}

#[test]
fn set_ior_union_update() {
    assert_eq!(
        run_python_one("s = {1, 2}\ns |= {2, 3}\nprint(sorted(s))\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn set_iand_intersection_update() {
    assert_eq!(
        run_python_one("s = {1, 2, 3}\ns &= {2, 3, 4}\nprint(sorted(s))\n"),
        "[2, 3]"
    );
}

#[test]
fn set_isub_difference_update() {
    assert_eq!(
        run_python_one("s = {1, 2, 3}\ns -= {2}\nprint(sorted(s))\n"),
        "[1, 3]"
    );
}

#[test]
fn set_ixor_symmetric_difference_update() {
    assert_eq!(
        run_python_one("s = {1, 2}\ns ^= {2, 3}\nprint(sorted(s))\n"),
        "[1, 3]"
    );
}

#[test]
fn dict_ior_merge() {
    assert_eq!(
        run_python_one("d = {'a': 1}\nd |= {'b': 2}\nprint(sorted(d.items()))\n"),
        "[('a', 1), ('b', 2)]"
    );
}

#[test]
fn list_iadd_tuple_source() {
    assert_eq!(
        run_python_one("xs = [1]\nxs += (2, 3)\nprint(xs)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn list_iadd_returns_none() {
    assert_eq!(
        run_python_one("xs = [1]\nr = xs.__iadd__([2])\nprint(r is xs)\n"),
        "True"
    );
}

#[test]
fn str_iadd_returns_same_object() {
    assert_eq!(
        run_python_one("s = 'a'\nt = s\ns += 'b'\nprint(s is t)\n"),
        "True"
    );
}

#[test]
fn augmented_assign_in_loop_counter() {
    assert_eq!(
        run_python_one("n = 0\nfor _ in range(4):\n n += 1\nprint(n)\n"),
        "4"
    );
}

#[test]
fn augmented_assign_on_indexed_list_item() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nxs[1] += 10\nprint(xs)\n"),
        "[1, 12, 3]"
    );
}

#[test]
fn augmented_assign_on_dict_value() {
    assert_eq!(
        run_python_one("d = {'n': 1}\nd['n'] += 4\nprint(d['n'])\n"),
        "5"
    );
}

#[test]
fn augmented_assign_bitwise_and_int() {
    assert_eq!(
        run_python_one("n = 15\nn &= 10\nprint(n)\n"),
        "10"
    );
}

#[test]
fn augmented_assign_bitwise_or_int() {
    assert_eq!(
        run_python_one("n = 1\nn |= 8\nprint(n)\n"),
        "9"
    );
}

#[test]
fn augmented_assign_bitwise_xor_int() {
    assert_eq!(
        run_python_one("n = 12\nn ^= 10\nprint(n)\n"),
        "6"
    );
}

#[test]
fn augmented_assign_lshift() {
    assert_eq!(
        run_python_one("n = 1\nn <<= 3\nprint(n)\n"),
        "8"
    );
}

#[test]
fn augmented_assign_rshift() {
    assert_eq!(
        run_python_one("n = 32\nn >>= 2\nprint(n)\n"),
        "8"
    );
}

#[test]
fn list_iadd_self_reference_safe() {
    assert_eq!(
        run_python_one("xs = [1]\nxs += xs\nprint(xs)\n"),
        "[1, 1]"
    );
}

#[test]
fn str_imul_one_unchanged() {
    assert_eq!(
        run_python_one("s = 'abc'\ns *= 1\nprint(s)\n"),
        "abc"
    );
}

#[test]
fn int_iadd_negative_delta() {
    assert_eq!(
        run_python_one("n = 5\nn += -2\nprint(n)\n"),
        "3"
    );
}

#[test]
fn float_imul_assign() {
    assert_eq!(
        run_python_one("x = 2.0\nx *= 1.5\nprint(x)\n"),
        "3.0"
    );
}

#[test]
fn list_imul_one_same_content() {
    assert_eq!(
        run_python_one("xs = [1, 2]\nxs *= 1\nprint(xs)\n"),
        "[1, 2]"
    );
}

#[test]
fn augmented_assign_chained_on_attribute() {
    assert_eq!(
        run_python_one("class C:\n def __init__(self):\n  self.n = 0\nc = C()\nc.n += 5\nc.n += 3\nprint(c.n)\n"),
        "8"
    );
}

#[test]
fn augmented_assign_on_nested_list_cell() {
    assert_eq!(
        run_python_one("grid = [[0]]\ngrid[0][0] += 7\nprint(grid)\n"),
        "[[7]]"
    );
}

#[test]
fn set_iadd_alias_for_ior() {
    assert_eq!(
        run_python_one("s = {1}\ns += {2}\nprint(sorted(s))\n"),
        "[1, 2]"
    );
}

#[test]
fn augmented_assign_expression_statement() {
    assert_eq!(
        run_python_one("n = 1\n(n := n)\nn += 1\nprint(n)\n"),
        "2"
    );
}
