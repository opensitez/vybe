use crate::helpers::run_python_one;

#[test]
fn fstring_interpolates_variable() {
    assert_eq!(run_python_one("x = 7\nprint(f'{x}')\n"), "7");
}

#[test]
fn fstring_interpolates_expression() {
    assert_eq!(run_python_one("print(f'{2 + 3}')\n"), "5");
}

#[test]
fn fstring_multiple_fields() {
    assert_eq!(run_python_one("a = 1\nb = 2\nprint(f'{a}-{b}')\n"), "1-2");
}

#[test]
fn fstring_format_spec_width() {
    assert_eq!(run_python_one("print(f'{3:03d}')\n"), "003");
}

#[test]
fn fstring_format_spec_float_precision() {
    assert_eq!(run_python_one("print(f'{1/3:.2f}')\n"), "0.33");
}

#[test]
fn fstring_format_spec_hex() {
    assert_eq!(run_python_one("print(f'{255:x}')\n"), "ff");
}

#[test]
fn fstring_format_spec_binary() {
    assert_eq!(run_python_one("print(f'{5:b}')\n"), "101");
}

#[test]
fn fstring_format_spec_percent_style() {
    assert_eq!(run_python_one("print(f'{0.5:.0%}')\n"), "50%");
}

#[test]
fn fstring_escape_braces_doubled() {
    assert_eq!(run_python_one("print(f'{{literal}} {1}')\n"), "literal 1");
}

#[test]
fn fstring_calls_method_on_value() {
    assert_eq!(run_python_one("print(f\"{'hi'.upper()}\")\n"), "HI");
}

#[test]
fn fstring_nested_quotes() {
    assert_eq!(run_python_one("print(f\"{'a'}\")\n"), "a");
}

#[test]
fn fstring_dict_access() {
    assert_eq!(
        run_python_one("d = {'k': 'v'}\nprint(f\"{d['k']}\")\n"),
        "v"
    );
}

#[test]
fn fstring_list_index() {
    assert_eq!(run_python_one("xs = [9]\nprint(f'{xs[0]}')\n"), "9");
}

#[test]
fn fstring_equals_debug_repr() {
    assert_eq!(run_python_one("print(f'{42!r}')\n"), "42");
}

#[test]
fn fstring_equals_str_conversion() {
    assert_eq!(run_python_one("print(f'{42!s}')\n"), "42");
}

#[test]
fn fstring_multiline_expression() {
    assert_eq!(run_python_one("x = 2\ny = 3\nprint(f'{x * y}')\n"), "6");
}

#[test]
fn fstring_with_comma_separator_format() {
    assert_eq!(run_python_one("print(f'{1000:,}')\n"), "1,000");
}

#[test]
fn fstring_align_right() {
    assert_eq!(run_python_one("print(f'{'hi':>5}')\n"), "   hi");
}

#[test]
fn fstring_align_left() {
    assert_eq!(run_python_one("print(f'{'hi':<5}')\n"), "hi   ");
}

#[test]
fn fstring_center_align() {
    assert_eq!(run_python_one("print(f'{'hi':^4}')\n"), " hi ");
}

#[test]
fn fstring_sign_plus_for_positive() {
    assert_eq!(run_python_one("print(f'{3:+d}')\n"), "+3");
}

#[test]
fn fstring_sign_space_for_positive() {
    assert_eq!(run_python_one("print(f'{3: d}')\n"), " 3");
}

#[test]
fn fstring_negative_int_preserved() {
    assert_eq!(run_python_one("print(f'{-4}')\n"), "-4");
}

#[test]
fn fstring_bool_value() {
    assert_eq!(run_python_one("print(f'{True}')\n"), "True");
}

#[test]
fn fstring_none_value() {
    assert_eq!(run_python_one("print(f'{None}')\n"), "None");
}

#[test]
fn fstring_concatenated_parts() {
    assert_eq!(
        run_python_one("name = 'py'\nprint(f'hello ' + f'{name}')\n"),
        "hello py"
    );
}

#[test]
fn fstring_in_return_from_function() {
    assert_eq!(
        run_python_one("def f():\n return f'val={1}'\nprint(f())\n"),
        "val=1"
    );
}

#[test]
fn fstring_nested_fstring() {
    assert_eq!(run_python_one("x = 1\nprint(f'{f\"{x}\"}')\n"), "1");
}

#[test]
fn fstring_with_percent_inside_expression() {
    assert_eq!(run_python_one("print(f'{10 % 3}')\n"), "1");
}

#[test]
fn fstring_datetime_like_manual() {
    assert_eq!(
        run_python_one("y, m, d = 2024, 5, 9\nprint(f'{y}-{m:02d}-{d:02d}')\n"),
        "2024-05-09"
    );
}

#[test]
fn fstring_json_like_fragment() {
    assert_eq!(
        run_python_one("k, v = 'a', 1\nprint(f'\"{k}\": {v}')\n"),
        "\"a\": 1"
    );
}

#[test]
fn fstring_equality_in_expression() {
    assert_eq!(
        run_python_one("print(f'{'yes' if 1 == 1 else 'no'}')\n"),
        "yes"
    );
}

#[test]
fn fstring_length_in_expression() {
    assert_eq!(run_python_one("s = 'abc'\nprint(f'{len(s)}')\n"), "3");
}

#[test]
fn fstring_strip_after_format() {
    assert_eq!(run_python_one("print(f'{'ab':>5}'.strip())\n"), "ab");
}

#[test]
fn fstring_float_scientific() {
    assert_eq!(run_python_one("print(f'{1234.5:e}')\n"), "1.234500e+03");
}

#[test]
fn fstring_int_with_underscores_display() {
    assert_eq!(run_python_one("n = 1_000_000\nprint(f'{n}')\n"), "1000000");
}

#[test]
fn fstring_tuple_display() {
    assert_eq!(run_python_one("print(f'{(1, 2)}')\n"), "(1, 2)");
}

#[test]
fn fstring_set_display_sorted() {
    assert_eq!(run_python_one("print(f'{sorted({1, 2})}')\n"), "[1, 2]");
}

#[test]
fn fstring_list_display() {
    assert_eq!(run_python_one("print(f'{[1, 2]}')\n"), "[1, 2]");
}

#[test]
fn fstring_backslash_in_literal_part() {
    assert_eq!(
        run_python_one("print(f'line\\nnext {1}')\n"),
        "line\nnext 1"
    );
}

#[test]
fn fstring_tab_in_literal() {
    assert_eq!(run_python_one("print(f'a\\tb')\n"), "a\tb");
}

#[test]
fn fstring_raw_prefix_not_with_f_mixed_use_regular() {
    assert_eq!(run_python_one("print(f'{'x'}')\n"), "x");
}

#[test]
fn fstring_calculation_precedence() {
    assert_eq!(run_python_one("print(f'{2 + 3 * 4}')\n"), "14");
}

#[test]
fn fstring_parenthesized_expression() {
    assert_eq!(run_python_one("print(f'{(2 + 3) * 4}')\n"), "20");
}

#[test]
fn fstring_multiple_format_specs() {
    assert_eq!(run_python_one("print(f'{1:>3}|{2:<3}')\n"), "  1|2  ");
}

#[test]
fn fstring_empty_braces_escape() {
    assert_eq!(run_python_one("print(f'{{}}')\n"), "{}");
}
