use crate::helpers::{run_print, run_python_one};

#[test]
fn walrus_in_if_condition() {
    assert_eq!(
        run_python_one("data = [1, 2, 3]\nif (n := len(data)) > 2:\n print(n)\n"),
        "3"
    );
}

#[test]
fn walrus_in_while_condition() {
    assert_eq!(
        run_python_one("it = iter([1, 2])\nwhile (x := next(it, None)) is not None:\n print(x)\n"),
        "1\n2"
    );
}

#[test]
fn walrus_in_comprehension_filter() {
    assert_eq!(
        run_print("[y for x in [1, 2, 3] if (y := x * 2) > 2]"),
        "[4, 6]"
    );
}

#[test]
fn walrus_assign_and_use_same_line() {
    assert_eq!(run_python_one("print((a := 5) + 1)\n"), "6");
}

#[test]
fn walrus_in_expression_list() {
    assert_eq!(run_python_one("print([(b := 2), b])\n"), "[2, 2]");
}

#[test]
fn walrus_scope_in_comprehension_local() {
    assert_eq!(
        run_python_one("out = [z for _ in [1] if (z := 9)]\nprint(z if 'z' in dir() else out)\n"),
        "[9]"
    );
}

#[test]
fn walrus_read_file_pattern_simulated() {
    assert_eq!(
        run_python_one("lines = ['a', '']\nif (line := lines[0]):\n print(line)\n"),
        "a"
    );
}

#[test]
fn walrus_regex_search_style() {
    assert_eq!(
        run_python_one("text = 'id=42'\nif (idx := text.find('=')) >= 0:\n print(text[idx+1:])\n"),
        "42"
    );
}

#[test]
fn walrus_sum_accumulator_pattern() {
    assert_eq!(
        run_python_one(
            "total = 0\nnums = [1, 2, 3]\nwhile nums and (total := total + nums.pop()):\n pass\nprint(total)\n"
        ),
        "6"
    );
}

#[test]
fn walrus_in_assert() {
    assert_eq!(
        run_python_one("x = 3\nassert (y := x * 2) == 6\nprint(y)\n"),
        "6"
    );
}

#[test]
fn walrus_nested_parentheses() {
    assert_eq!(run_python_one("print(((v := 10)))\n"), "10");
}

#[test]
fn walrus_in_fstring_not_allowed_use_prior() {
    assert_eq!(run_python_one("(n := 4)\nprint(f'{n}')\n"), "4");
}

#[test]
fn walrus_list_comp_value_reuse() {
    assert_eq!(run_print("[(s := str(i)) for i in range(2)]"), "['0', '1']");
}

#[test]
fn walrus_dict_comp() {
    assert_eq!(
        run_print("{k: (v := k * 2) for k in range(2)}"),
        "{0: 0, 1: 2}"
    );
}

#[test]
fn walrus_set_comp() {
    assert_eq!(
        run_print("sorted({(y := x + 1) for x in range(2)})"),
        "[1, 2]"
    );
}

#[test]
fn walrus_generator_expression() {
    assert_eq!(
        run_python_one("print(sum((n := x) for x in range(3)))\n"),
        "3"
    );
}

#[test]
fn walrus_multiple_in_tuple_unpack() {
    assert_eq!(run_python_one("print((a := 1, b := 2))\n"), "(1, 2)");
}

#[test]
fn walrus_in_function_body() {
    assert_eq!(
        run_python_one("def f():\n if (x := 5) > 0:\n  return x\nprint(f())\n"),
        "5"
    );
}

#[test]
fn walrus_truthy_check() {
    assert_eq!(
        run_python_one("items = ['x']\nif item := items[0]:\n print(item)\n"),
        "x"
    );
}

#[test]
fn walrus_falsy_skips_block() {
    assert_eq!(
        run_python_one("items = ['']\nif item := items[0]:\n print('yes')\nelse:\n print('no')\n"),
        "no"
    );
}

#[test]
fn walrus_in_elif_chain() {
    assert_eq!(
        run_python_one("x = 5\nif False:\n pass\nelif (n := x - 2) > 2:\n print(n)\n"),
        "3"
    );
}

#[test]
fn walrus_string_processing() {
    assert_eq!(
        run_python_one("parts = 'a,b'.split(',')\nif (first := parts[0]) == 'a':\n print(first)\n"),
        "a"
    );
}

#[test]
fn walrus_min_max_pattern() {
    assert_eq!(
        run_python_one("vals = [3, 1, 2]\nprint((m := min(vals)) + max(vals))\n"),
        "5"
    );
}

#[test]
fn walrus_list_get_with_default() {
    assert_eq!(
        run_python_one("d = {1: 'a'}\nprint((v := d.get(2, 'missing')))\n"),
        "missing"
    );
}

#[test]
fn walrus_reuse_in_same_expression() {
    assert_eq!(run_python_one("print((k := 3) + k)\n"), "6");
}

#[test]
fn walrus_in_boolean_and_short_circuit() {
    assert_eq!(
        run_python_one("flag = True\nprint(flag and (n := 7))\n"),
        "7"
    );
}

#[test]
fn walrus_in_boolean_or() {
    assert_eq!(
        run_python_one("flag = False\nprint(flag or (n := 8))\n"),
        "8"
    );
}

#[test]
fn walrus_for_loop_read() {
    assert_eq!(
        run_python_one(
            "pairs = [('a', 1)]\nfor k, v in pairs:\n if (label := k + str(v)):\n  print(label)\n"
        ),
        "a1"
    );
}

#[test]
fn walrus_while_read_lines_style() {
    assert_eq!(
        run_python_one(
            "lines = iter(['x', ''])\nout = []\nwhile (line := next(lines, '')):\n out.append(line)\nprint(out)\n"
        ),
        "['x']"
    );
}

#[test]
fn walrus_match_guard_style_manual() {
    assert_eq!(
        run_python_one("x = 4\nif (half := x // 2) == 2:\n print(half)\n"),
        "2"
    );
}

#[test]
fn walrus_assign_none() {
    assert_eq!(run_python_one("print((x := None) is None)\n"), "True");
}

#[test]
fn walrus_assign_list() {
    assert_eq!(run_python_one("print(len(xs := [1, 2, 3]))\n"), "3");
}

#[test]
fn walrus_assign_dict() {
    assert_eq!(run_python_one("print((d := {'a': 1})['a'])\n"), "1");
}

#[test]
fn walrus_in_try_block() {
    assert_eq!(
        run_python_one("try:\n print((n := 2) + 1)\nexcept:\n pass\n"),
        "3"
    );
}

#[test]
fn walrus_comprehension_nested() {
    assert_eq!(run_print("[(a := x) + 1 for x in range(2)]"), "[1, 2]");
}

#[test]
fn walrus_negative_number() {
    assert_eq!(run_python_one("print((n := -5) + 10)\n"), "5");
}

#[test]
fn walrus_float_value() {
    assert_eq!(run_python_one("print((f := 2.5) * 2)\n"), "5.0");
}

#[test]
fn walrus_string_concat() {
    assert_eq!(run_python_one("print((s := 'a') + 'b')\n"), "ab");
}

#[test]
fn walrus_identity_check() {
    assert_eq!(
        run_python_one("obj = []\nprint((same := obj) is obj)\n"),
        "True"
    );
}

#[test]
fn walrus_membership_result() {
    assert_eq!(run_python_one("print((found := 2 in [1, 2, 3]))\n"), "True");
}

#[test]
fn walrus_type_check() {
    assert_eq!(
        run_python_one("print((is_str := isinstance('a', str)))\n"),
        "True"
    );
}

#[test]
fn walrus_len_in_condition() {
    assert_eq!(
        run_python_one("data = 'abc'\nif (ln := len(data)) == 3:\n print(ln)\n"),
        "3"
    );
}

#[test]
fn walrus_pow_computation() {
    assert_eq!(run_python_one("print((p := 2 ** 3))\n"), "8");
}

#[test]
fn walrus_modulo_computation() {
    assert_eq!(run_python_one("print((r := 10 % 3))\n"), "1");
}

#[test]
fn walrus_floor_div() {
    assert_eq!(run_python_one("print((q := 7 // 2))\n"), "3");
}
