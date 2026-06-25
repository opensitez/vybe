use crate::helpers::{run_python_one, run_print};

#[test]
fn unpack_tuple_to_variables() {
    assert_eq!(run_python_one("a, b = (1, 2)\nprint(a, b)\n"), "1 2");
}

#[test]
fn unpack_list_to_variables() {
    assert_eq!(run_python_one("x, y, z = [1, 2, 3]\nprint(z)\n"), "3");
}

#[test]
fn unpack_nested_tuple() {
    assert_eq!(
        run_python_one("((a, b), c) = ((1, 2), 3)\nprint(a, b, c)\n"),
        "1 2 3"
    );
}

#[test]
fn unpack_nested_list_in_tuple() {
    assert_eq!(
        run_python_one("(a, [b, c]) = (0, [1, 2])\nprint(b, c)\n"),
        "1 2"
    );
}

#[test]
fn unpack_star_collects_middle() {
    assert_eq!(
        run_python_one("first, *rest = [1, 2, 3, 4]\nprint(first, rest)\n"),
        "1 [2, 3, 4]"
    );
}

#[test]
fn unpack_star_at_end() {
    assert_eq!(
        run_python_one("a, b, *tail = [1, 2, 3]\nprint(tail)\n"),
        "[3]"
    );
}

#[test]
fn unpack_star_only() {
    assert_eq!(
        run_python_one("*all, = [1, 2]\nprint(all)\n"),
        "[1, 2]"
    );
}

#[test]
fn unpack_in_for_loop() {
    assert_eq!(
        run_python_one("total = 0\nfor a, b in [(1, 2), (3, 4)]:\n total += a + b\nprint(total)\n"),
        "10"
    );
}

#[test]
fn unpack_enumerate() {
    assert_eq!(
        run_python_one("for i, v in enumerate(['a']):\n print(i, v)\n"),
        "0 a"
    );
}

#[test]
fn unpack_zip_three() {
    assert_eq!(
        run_python_one("for a, b, c in zip([1], [2], [3]):\n print(a, b, c)\n"),
        "1 2 3"
    );
}

#[test]
fn unpack_dict_items() {
    assert_eq!(
        run_python_one("d = {'a': 1}\nfor k, v in d.items():\n print(k, v)\n"),
        "a 1"
    );
}

#[test]
fn unpack_dict_keys_values() {
    assert_eq!(
        run_python_one("d = {'x': 9}\nprint(list(d.keys())[0], list(d.values())[0])\n"),
        "x 9"
    );
}

#[test]
fn unpack_swap_variables() {
    assert_eq!(
        run_python_one("a, b = 1, 2\na, b = b, a\nprint(a, b)\n"),
        "2 1"
    );
}

#[test]
fn unpack_function_return_tuple() {
    assert_eq!(
        run_python_one("def pair():\n return 4, 5\nx, y = pair()\nprint(x + y)\n"),
        "9"
    );
}

#[test]
fn unpack_ignores_with_underscore() {
    assert_eq!(
        run_python_one("a, _, c = (1, 2, 3)\nprint(a, c)\n"),
        "1 3"
    );
}

#[test]
fn unpack_extended_iterable_unrolling() {
    assert_eq!(
        run_print("[*range(3)]"),
        "[0, 1, 2]"
    );
}

#[test]
fn unpack_merge_lists() {
    assert_eq!(
        run_print("[*[1, 2], *[3]]"),
        "[1, 2, 3]"
    );
}

#[test]
fn unpack_in_list_literal_middle() {
    assert_eq!(
        run_print("[0, *range(2), 9]"),
        "[0, 0, 1, 9]"
    );
}

#[test]
fn unpack_in_tuple_literal() {
    assert_eq!(
        run_print("(*(1, 2), 3)"),
        "(1, 2, 3)"
    );
}

#[test]
fn unpack_in_set_literal() {
    assert_eq!(
        run_print("sorted({*'ab', *'bc'})"),
        "['a', 'b', 'c']"
    );
}

#[test]
fn unpack_string_to_list() {
    assert_eq!(run_print("[*'hi']"), "['h', 'i']");
}

#[test]
fn unpack_dict_merge_literals() {
    assert_eq!(
        run_print("{**{'a': 1}, **{'b': 2}}"),
        "{'a': 1, 'b': 2}"
    );
}

#[test]
fn unpack_dict_merge_override() {
    assert_eq!(
        run_print("{**{'a': 1}, **{'a': 2}}"),
        "{'a': 2}"
    );
}

#[test]
fn unpack_call_with_star_args() {
    assert_eq!(
        run_python_one("def f(a, b):\n return a + b\nprint(f(*[2, 3]))\n"),
        "5"
    );
}

#[test]
fn unpack_call_with_starstar_kwargs() {
    assert_eq!(
        run_python_one("def f(x, y):\n return x * y\nprint(f(**{'x': 3, 'y': 4}))\n"),
        "12"
    );
}

#[test]
fn unpack_mixed_positional_and_keyword() {
    assert_eq!(
        run_python_one("def f(a, b, c=0):\n return a + b + c\nprint(f(1, *[2], **{'c': 3}))\n"),
        "6"
    );
}

#[test]
fn unpack_nested_star_in_list() {
    assert_eq!(
        run_print("[*[1, 2], *[3, *[4]]]"),
        "[1, 2, 3, 4]"
    );
}

#[test]
fn unpack_from_generator() {
    assert_eq!(
        run_print("[*(x for x in range(2))]"),
        "[0, 1]"
    );
}

#[test]
fn unpack_tuple_on_rhs_from_list() {
    assert_eq!(
        run_python_one("t = (1, 2)\na, b = list(t)\nprint(a, b)\n"),
        "1 2"
    );
}

#[test]
fn unpack_length_mismatch_raises() {
    assert_eq!(
        run_python_one("try:\n a, b = [1]\nexcept ValueError:\n print('val')\n"),
        "val"
    );
}

#[test]
fn unpack_too_many_values_raises() {
    assert_eq!(
        run_python_one("try:\n a, = [1, 2]\nexcept ValueError:\n print('val')\n"),
        "val"
    );
}

#[test]
fn unpack_slice_assign_from_list() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\na[:2] = ['x', 'y']\nprint(a)\n"),
        "['x', 'y', 3]"
    );
}

#[test]
fn unpack_parallel_assignment_chain() {
    assert_eq!(
        run_python_one("a = b = c = 0\na, b = 1, 2\nprint(a, b, c)\n"),
        "1 2 0"
    );
}

#[test]
fn unpack_from_os_walk_style_pairs() {
    assert_eq!(
        run_python_one("pairs = [('a', 1), ('b', 2)]\nout = []\nfor k, v in pairs:\n out.append(k + str(v))\nprint(out)\n"),
        "['a1', 'b2']"
    );
}

#[test]
fn unpack_matrix_rows() {
    assert_eq!(
        run_python_one("rows = [(1, 2), (3, 4)]\ns = 0\nfor x, y in rows:\n s += x + y\nprint(s)\n"),
        "10"
    );
}

#[test]
fn unpack_with_rest_and_tail() {
    assert_eq!(
        run_python_one("a, *mid, z = range(5)\nprint(a, len(mid), z)\n"),
        "0 3 4"
    );
}

#[test]
fn unpack_multiple_stars_in_call() {
    assert_eq!(
        run_python_one("def f(*args):\n return sum(args)\nprint(f(*[1, 2], *[3]))\n"),
        "6"
    );
}

#[test]
fn unpack_dict_in_function_params() {
    assert_eq!(
        run_python_one("def f(**kw):\n return kw\nprint(f(**{'a': 1})['a'])\n"),
        "1"
    );
}

#[test]
fn unpack_list_comp_with_star() {
    assert_eq!(
        run_print("[*['a'], *['b']]"),
        "['a', 'b']"
    );
}

#[test]
fn unpack_tuple_of_lists_flatten() {
    assert_eq!(
        run_print("[*([1, 2]), *([3])]"),
        "[1, 2, 3]"
    );
}

#[test]
fn unpack_assign_from_split() {
    assert_eq!(
        run_python_one("a, b = 'x,y'.split(',')\nprint(a, b)\n"),
        "x y"
    );
}

#[test]
fn unpack_head_tail_recursive_style() {
    assert_eq!(
        run_python_one("def sum_list(xs):\n if not xs:\n  return 0\n head, *tail = xs\n return head + sum_list(tail)\nprint(sum_list([1, 2, 3]))\n"),
        "6"
    );
}

#[test]
fn unpack_from_map_result() {
    assert_eq!(
        run_python_one("a, b = map(int, ['2', '3'])\nprint(a + b)\n"),
        "5"
    );
}

#[test]
fn unpack_from_zip_single_row() {
    assert_eq!(
        run_python_one("names, ages = zip(('Ann', 20))\nprint(names[0], ages[0])\n"),
        "Ann 20"
    );
}

#[test]
fn unpack_empty_star_list() {
    assert_eq!(
        run_python_one("a, *rest = [1]\nprint(rest)\n"),
        "[]"
    );
}
