use crate::helpers::{run_print, run_python_one};

#[test]
fn filter_none_removes_falsy_zero() {
    assert_eq!(
        run_print("list(filter(None, [0, 1, 2, '', 3]))"),
        "[1, 2, 3]"
    );
}

#[test]
fn filter_lambda_even_only() {
    assert_eq!(
        run_print("list(filter(lambda x: x % 2 == 0, range(6)))"),
        "[0, 2, 4]"
    );
}

#[test]
fn filter_lambda_positive_only() {
    assert_eq!(
        run_print("list(filter(lambda x: x > 0, [-1, 0, 2, 3]))"),
        "[2, 3]"
    );
}

#[test]
fn map_lambda_square() {
    assert_eq!(
        run_print("list(map(lambda x: x * x, [1, 2, 3]))"),
        "[1, 4, 9]"
    );
}

#[test]
fn map_builtin_str_to_strings() {
    assert_eq!(run_print("list(map(str, [1, 2, 3]))"), "['1', '2', '3']");
}

#[test]
fn map_two_iterables_add_pairs() {
    assert_eq!(
        run_print("list(map(lambda a, b: a + b, [1, 2], [10, 20]))"),
        "[11, 22]"
    );
}

#[test]
fn filter_empty_input() {
    assert_eq!(run_print("list(filter(None, []))"), "[]");
}

#[test]
fn map_empty_input() {
    assert_eq!(run_print("list(map(str, []))"), "[]");
}

#[test]
fn any_on_filter_nonzero() {
    assert_eq!(run_print("any(x > 0 for x in [-1, 0, 2])"), "True");
}

#[test]
fn all_on_map_positive() {
    assert_eq!(run_print("all(x > 0 for x in map(abs, [-1, -2]))"), "True");
}

#[test]
fn filter_with_list_comprehension_equivalent() {
    assert_eq!(
        run_python_one(
            "a = list(filter(lambda x: x % 2 == 1, range(5)))\nb = [x for x in range(5) if x % 2 == 1]\nprint(a == b)\n"
        ),
        "True"
    );
}

#[test]
fn map_lazy_consumed_once() {
    assert_eq!(
        run_python_one("it = map(lambda x: x + 1, [1, 2])\nprint(list(it))\n"),
        "[2, 3]"
    );
}

#[test]
fn filter_on_string_chars_alnum() {
    assert_eq!(
        run_print("list(filter(str.isdigit, 'a1b2c3'))"),
        "['1', '2', '3']"
    );
}

#[test]
fn map_chr_from_ascii_codes() {
    assert_eq!(run_print("list(map(chr, [65, 66, 67]))"), "['A', 'B', 'C']");
}

#[test]
fn filter_none_on_list_with_false_and_true() {
    assert_eq!(
        run_print("list(filter(None, [False, True, 0, 1]))"),
        "[True, 1]"
    );
}

#[test]
fn map_len_on_words() {
    assert_eq!(run_print("list(map(len, ['hi', 'hello']))"), "[2, 5]");
}

#[test]
fn any_empty_generator_false() {
    assert_eq!(run_print("any(x > 0 for x in [])"), "False");
}

#[test]
fn all_empty_generator_true() {
    assert_eq!(run_print("all(x > 0 for x in [])"), "True");
}

#[test]
fn all_short_circuit_on_first_false() {
    assert_eq!(
        run_python_one("def gen():\n yield True\n yield False\n yield True\nprint(all(gen()))\n"),
        "False"
    );
}

#[test]
fn any_short_circuit_on_first_true() {
    assert_eq!(
        run_python_one("def gen():\n yield False\n yield True\n yield False\nprint(any(gen()))\n"),
        "True"
    );
}

#[test]
fn map_none_function_identity() {
    assert_eq!(run_print("list(map(lambda x: x, [1, 2, 3]))"), "[1, 2, 3]");
}

#[test]
fn filter_identity_lambda_keeps_all_truthy() {
    assert_eq!(
        run_print("list(filter(lambda x: x, [1, 2, 3]))"),
        "[1, 2, 3]"
    );
}

#[test]
fn map_with_float_conversion() {
    assert_eq!(run_print("list(map(float, ['1.5', '2.5']))"), "[1.5, 2.5]");
}

#[test]
fn filter_strings_nonempty() {
    assert_eq!(
        run_print("list(filter(None, ['', 'a', '', 'b']))"),
        "['a', 'b']"
    );
}

#[test]
fn map_ord_on_characters() {
    assert_eq!(run_print("list(map(ord, 'AB'))"), "[65, 66]");
}

#[test]
fn any_on_map_even_exists() {
    assert_eq!(
        run_print("any(x % 2 == 0 for x in map(int, ['1', '2', '3']))"),
        "True"
    );
}

#[test]
fn all_on_range_small() {
    assert_eq!(run_print("all(x >= 0 for x in range(5))"), "True");
}

#[test]
fn filter_dict_items_by_value() {
    assert_eq!(
        run_print("list(filter(lambda kv: kv[1] > 1, {'a': 1, 'b': 2}.items()))"),
        "[('b', 2)]"
    );
}

#[test]
fn map_extract_dict_keys() {
    assert_eq!(
        run_print("list(map(lambda kv: kv[0], {'x': 1, 'y': 2}.items()))"),
        "['x', 'y']"
    );
}

#[test]
fn filter_on_tuple_of_mixed_types() {
    assert_eq!(
        run_print("list(filter(lambda x: isinstance(x, int), (1, 'a', 2)))"),
        "[1, 2]"
    );
}

#[test]
fn map_three_arg_zip_style_manual() {
    assert_eq!(
        run_print("list(map(lambda a, b, c: a + b + c, [1], [2], [3]))"),
        "[6]"
    );
}

#[test]
fn any_with_bool_map() {
    assert_eq!(run_print("any(map(bool, [0, 0, 1]))"), "True");
}

#[test]
fn all_with_bool_map_all_true() {
    assert_eq!(run_print("all(map(bool, [1, 2, 3]))"), "True");
}

#[test]
fn filter_lambda_none_explicitly_false() {
    assert_eq!(
        run_print("list(filter(lambda x: x is not None, [None, 1, None, 2]))"),
        "[1, 2]"
    );
}

#[test]
fn map_to_bool_list() {
    assert_eq!(
        run_print("list(map(bool, [0, 1, 2]))"),
        "[False, True, True]"
    );
}

#[test]
fn filter_on_range_stop_at_five() {
    assert_eq!(
        run_print("list(filter(lambda x: x < 3, range(5)))"),
        "[0, 1, 2]"
    );
}

#[test]
fn map_increment_strings_with_suffix() {
    assert_eq!(
        run_print("list(map(lambda s: s + '!', ['a', 'b']))"),
        "['a!', 'b!']"
    );
}

#[test]
fn any_on_nested_comprehension() {
    assert_eq!(
        run_print("any(x > 3 for row in [[1, 2], [4, 5]] for x in row)"),
        "True"
    );
}

#[test]
fn all_elements_equal_via_all() {
    assert_eq!(run_print("all(x == 2 for x in [2, 2, 2])"), "True");
}

#[test]
fn filter_map_chain_square_evens() {
    assert_eq!(
        run_print("list(map(lambda x: x * x, filter(lambda x: x % 2 == 0, range(5))))"),
        "[0, 4, 16]"
    );
}
