use crate::helpers::run_python_one;

#[test]
fn augmented_add_int() {
    assert_eq!(run_python_one("x = 1\nx += 2\nprint(x)\n"), "3");
}

#[test]
fn augmented_sub_int() {
    assert_eq!(run_python_one("x = 5\nx -= 2\nprint(x)\n"), "3");
}

#[test]
fn augmented_mul_int() {
    assert_eq!(run_python_one("x = 3\nx *= 4\nprint(x)\n"), "12");
}

#[test]
fn augmented_div_float() {
    assert_eq!(run_python_one("x = 7\nx /= 2\nprint(x)\n"), "3.5");
}

#[test]
fn augmented_floordiv() {
    assert_eq!(run_python_one("x = 7\nx //= 2\nprint(x)\n"), "3");
}

#[test]
fn augmented_mod() {
    assert_eq!(run_python_one("x = 10\nx %= 3\nprint(x)\n"), "1");
}

#[test]
fn augmented_pow() {
    assert_eq!(run_python_one("x = 2\nx **= 3\nprint(x)\n"), "8");
}

#[test]
fn augmented_add_string_concat() {
    assert_eq!(run_python_one("s = 'a'\ns += 'b'\nprint(s)\n"), "ab");
}

#[test]
fn augmented_mul_string_repeat() {
    assert_eq!(run_python_one("s = 'x'\ns *= 3\nprint(s)\n"), "xxx");
}

#[test]
fn augmented_add_list_extend_like() {
    assert_eq!(
        run_python_one("a = [1]\na += [2, 3]\nprint(a)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn augmented_mul_list_repeat() {
    assert_eq!(run_python_one("a = [1]\na *= 2\nprint(a)\n"), "[1, 1]");
}

#[test]
fn augmented_add_set_union() {
    assert_eq!(
        run_python_one("s = {1}\ns |= {2}\nprint(sorted(s))\n"),
        "[1, 2]"
    );
}

#[test]
fn augmented_and_set_intersection() {
    assert_eq!(
        run_python_one("s = {1, 2}\ns &= {2, 3}\nprint(sorted(s))\n"),
        "[2]"
    );
}

#[test]
fn augmented_xor_set_symmetric() {
    assert_eq!(
        run_python_one("s = {1, 2}\ns ^= {2, 3}\nprint(sorted(s))\n"),
        "[1, 3]"
    );
}

#[test]
fn augmented_sub_set_difference() {
    assert_eq!(
        run_python_one("s = {1, 2, 3}\ns -= {2}\nprint(sorted(s))\n"),
        "[1, 3]"
    );
}

#[test]
fn augmented_or_dict_merge() {
    assert_eq!(
        run_python_one("d = {'a': 1}\nd |= {'b': 2}\nprint(d)\n"),
        "{'a': 1, 'b': 2}"
    );
}

#[test]
fn augmented_add_float() {
    assert_eq!(run_python_one("x = 0.5\nx += 0.25\nprint(x)\n"), "0.75");
}

#[test]
fn augmented_sub_float() {
    assert_eq!(run_python_one("x = 1.0\nx -= 0.25\nprint(x)\n"), "0.75");
}

#[test]
fn augmented_mul_float() {
    assert_eq!(run_python_one("x = 2.0\nx *= 1.5\nprint(x)\n"), "3.0");
}

#[test]
fn augmented_div_assign_int_to_float() {
    assert_eq!(run_python_one("x = 4\nx /= 2\nprint(x)\n"), "2.0");
}

#[test]
fn augmented_add_tuple_creates_new() {
    assert_eq!(run_python_one("t = (1,)\nt += (2,)\nprint(t)\n"), "(1, 2)");
}

#[test]
fn augmented_add_bytes_concat() {
    assert_eq!(run_python_one("b = b'a'\nb += b'b'\nprint(b)\n"), "b'ab'");
}

#[test]
fn augmented_mul_bytes_repeat() {
    assert_eq!(run_python_one("b = b'x'\nb *= 2\nprint(b)\n"), "b'xx'");
}

#[test]
fn augmented_add_bytearray() {
    assert_eq!(
        run_python_one("ba = bytearray(b'a')\nba += b'b'\nprint(ba)\n"),
        "bytearray(b'ab')"
    );
}

#[test]
fn augmented_sub_list_not_supported() {
    assert_eq!(
        run_python_one("try:\n a = [1, 2]\n a -= [1]\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}

#[test]
fn augmented_mod_on_negative() {
    assert_eq!(run_python_one("x = -10\nx %= 3\nprint(x)\n"), "2");
}

#[test]
fn augmented_floordiv_negative() {
    assert_eq!(run_python_one("x = -7\nx //= 2\nprint(x)\n"), "-4");
}

#[test]
fn augmented_pow_zero_exp() {
    assert_eq!(run_python_one("x = 99\nx **= 0\nprint(x)\n"), "1");
}

#[test]
fn augmented_add_on_attr() {
    assert_eq!(
        run_python_one(
            "class C:\n def __init__(self):\n  self.n = 1\nc = C()\nc.n += 2\nprint(c.n)\n"
        ),
        "3"
    );
}

#[test]
fn augmented_add_on_subscript() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\na[1] += 10\nprint(a)\n"),
        "[1, 12, 3]"
    );
}

#[test]
fn augmented_mul_on_subscript() {
    assert_eq!(
        run_python_one("a = [2, 3]\na[0] *= 5\nprint(a)\n"),
        "[10, 3]"
    );
}

#[test]
fn augmented_add_dict_key_accumulate() {
    assert_eq!(
        run_python_one("d = {'n': 1}\nd['n'] += 4\nprint(d['n'])\n"),
        "5"
    );
}

#[test]
fn augmented_concat_in_loop() {
    assert_eq!(
        run_python_one("s = ''\nfor ch in 'ab':\n s += ch\nprint(s)\n"),
        "ab"
    );
}

#[test]
fn augmented_sum_in_loop() {
    assert_eq!(
        run_python_one("total = 0\nfor i in range(4):\n total += i\nprint(total)\n"),
        "6"
    );
}

#[test]
fn augmented_mul_accumulate_powers() {
    assert_eq!(
        run_python_one("p = 1\nfor _ in range(3):\n p *= 2\nprint(p)\n"),
        "8"
    );
}

#[test]
fn augmented_and_on_bool_int() {
    assert_eq!(run_python_one("x = 15\nx &= 7\nprint(x)\n"), "7");
}

#[test]
fn augmented_or_on_bool_int() {
    assert_eq!(run_python_one("x = 1\nx |= 8\nprint(x)\n"), "9");
}

#[test]
fn augmented_xor_on_bool_int() {
    assert_eq!(run_python_one("x = 6\nx ^= 3\nprint(x)\n"), "5");
}

#[test]
fn augmented_lshift_assign() {
    assert_eq!(run_python_one("x = 1\nx <<= 3\nprint(x)\n"), "8");
}

#[test]
fn augmented_rshift_assign() {
    assert_eq!(run_python_one("x = 32\nx >>= 3\nprint(x)\n"), "4");
}

#[test]
fn augmented_add_complex() {
    assert_eq!(
        run_python_one("z = 1 + 2j\nz += 3 + 4j\nprint(z)\n"),
        "(4+6j)"
    );
}

#[test]
fn augmented_sub_complex() {
    assert_eq!(
        run_python_one("z = 5 + 6j\nz -= 1 + 1j\nprint(z)\n"),
        "(4+5j)"
    );
}

#[test]
fn augmented_mul_complex() {
    assert_eq!(run_python_one("z = 1 + 1j\nz *= 2\nprint(z)\n"), "(2+2j)");
}

#[test]
fn augmented_div_complex() {
    assert_eq!(run_python_one("z = 4 + 4j\nz /= 2\nprint(z)\n"), "(2+2j)");
}

#[test]
fn augmented_add_none_raises() {
    assert_eq!(
        run_python_one("try:\n x = None\n x += 1\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}
