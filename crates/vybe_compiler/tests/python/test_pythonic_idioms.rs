use crate::helpers::{run_print, run_python_one};

#[test]
fn ternary_conditional_expression_basic() {
    assert_eq!(run_print("'big' if 10 > 5 else 'small'"), "big");
}

#[test]
fn ternary_nested_in_expression() {
    assert_eq!(
        run_python_one("x = 0\nprint('zero' if x == 0 else ('pos' if x > 0 else 'neg'))\n"),
        "zero"
    );
}

#[test]
fn ternary_in_list_comprehension() {
    assert_eq!(
        run_print("[('even' if i % 2 == 0 else 'odd') for i in range(3)]"),
        "['even', 'odd', 'even']"
    );
}

#[test]
fn ternary_in_dict_comprehension() {
    assert_eq!(
        run_print("{i: ('pos' if i > 0 else 'nonpos') for i in [-1, 0, 1]}"),
        "{-1: 'nonpos', 0: 'nonpos', 1: 'pos'}"
    );
}

#[test]
fn walrus_in_if_condition() {
    assert_eq!(
        run_python_one("data = [1, 2, 3]\nif (n := len(data)) > 2:\n print(n)\n"),
        "3"
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
fn walrus_while_read_lines_pattern() {
    assert_eq!(
        run_python_one(
            "lines = iter(['a', ''])\nout = []\nwhile (line := next(lines, '')):\n out.append(line)\nprint(out)\n"
        ),
        "['a']"
    );
}

#[test]
fn or_default_empty_string() {
    assert_eq!(run_print("'' or 'default'"), "default");
}

#[test]
fn or_default_zero_int() {
    assert_eq!(run_print("0 or 42"), "42");
}

#[test]
fn or_default_none() {
    assert_eq!(run_print("None or 'fallback'"), "fallback");
}

#[test]
fn and_short_circuit_false() {
    assert_eq!(run_print("0 and 99"), "0");
}

#[test]
fn and_short_circuit_true_returns_last() {
    assert_eq!(run_print("1 and 2 and 3"), "3");
}

#[test]
fn chained_comparison_in_if() {
    assert_eq!(
        run_python_one("x = 5\nprint('ok' if 0 < x < 10 else 'no')\n"),
        "ok"
    );
}

#[test]
fn chained_comparison_false() {
    assert_eq!(run_print("5 < 3 < 10"), "False");
}

#[test]
fn list_comp_double_filter() {
    assert_eq!(
        run_print("[x for x in range(10) if x % 2 == 0 if x % 3 == 0]"),
        "[0, 6]"
    );
}

#[test]
fn list_comp_nested_loops() {
    assert_eq!(
        run_print("[(i, j) for i in range(2) for j in range(2)]"),
        "[(0, 0), (0, 1), (1, 0), (1, 1)]"
    );
}

#[test]
fn dict_comp_with_if_filter() {
    assert_eq!(
        run_print("{k: v for k, v in [('a', 1), ('b', 2)] if v > 1}"),
        "{'b': 2}"
    );
}

#[test]
fn set_comp_with_transform() {
    assert_eq!(
        run_print("sorted({x * x for x in range(4)})"),
        "[0, 1, 4, 9]"
    );
}

#[test]
fn generator_expr_in_sum() {
    assert_eq!(
        run_python_one("print(sum(x * x for x in range(4)))\n"),
        "14"
    );
}

#[test]
fn generator_expr_in_any() {
    assert_eq!(
        run_python_one("print(any(x > 2 for x in [1, 2, 3]))\n"),
        "True"
    );
}

#[test]
fn generator_expr_in_all() {
    assert_eq!(
        run_python_one("print(all(x > 0 for x in [1, 2, 3]))\n"),
        "True"
    );
}

#[test]
fn generator_expr_in_max() {
    assert_eq!(
        run_python_one("print(max(len(s) for s in ['a', 'bbb']))\n"),
        "3"
    );
}

#[test]
fn generator_expr_in_min() {
    assert_eq!(run_python_one("print(min(x for x in [3, 1, 2]))\n"), "1");
}

#[test]
fn dict_merge_pipe_operator() {
    assert_eq!(run_print("{'a': 1} | {'b': 2}"), "{'a': 1, 'b': 2}");
}

#[test]
fn dict_merge_pipe_override() {
    assert_eq!(run_print("{'a': 1} | {'a': 2}"), "{'a': 2}");
}

#[test]
fn starred_list_literal_merge() {
    assert_eq!(run_print("[*[1, 2], *[3, 4]]"), "[1, 2, 3, 4]");
}

#[test]
fn starred_dict_literal_merge() {
    assert_eq!(run_print("{**{'a': 1}, **{'b': 2}}"), "{'a': 1, 'b': 2}");
}

#[test]
fn starred_call_unpack_list() {
    assert_eq!(
        run_python_one("def f(a, b):\n return a + b\nprint(f(*[2, 3]))\n"),
        "5"
    );
}

#[test]
fn double_star_call_unpack_dict() {
    assert_eq!(
        run_python_one("def f(x, y):\n return x * y\nprint(f(**{'x': 3, 'y': 4}))\n"),
        "12"
    );
}

#[test]
fn chained_assignment_integers() {
    assert_eq!(run_python_one("a = b = c = 0\na = 1\nprint(b, c)\n"), "0 0");
}

#[test]
fn parallel_unpack_swap() {
    assert_eq!(
        run_python_one("a, b = 1, 2\na, b = b, a\nprint(a, b)\n"),
        "2 1"
    );
}

#[test]
fn slice_assignment_replace_middle() {
    assert_eq!(
        run_python_one("a = [1, 2, 3, 4]\na[1:3] = [9]\nprint(a)\n"),
        "[1, 9, 4]"
    );
}

#[test]
fn slice_assignment_delete_range() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\ndel a[0:2]\nprint(a)\n"),
        "[3]"
    );
}

#[test]
fn enumerate_in_for_loop() {
    assert_eq!(
        run_python_one("for i, v in enumerate(['a', 'b']):\n if i == 1:\n  print(v)\n"),
        "b"
    );
}

#[test]
fn zip_in_for_loop() {
    assert_eq!(
        run_python_one("for a, b in zip([1, 2], [10, 20]):\n print(a + b)\n"),
        "11\n22"
    );
}

#[test]
fn sorted_with_key_lambda() {
    assert_eq!(
        run_print("sorted(['bb', 'a', 'ccc'], key=lambda s: len(s))"),
        "['a', 'bb', 'ccc']"
    );
}

#[test]
fn sorted_reverse_true() {
    assert_eq!(run_print("sorted([3, 1, 2], reverse=True)"), "[3, 2, 1]");
}

#[test]
fn filter_none_removes_falsy() {
    assert_eq!(run_print("list(filter(None, [0, 1, '', 'x']))"), "[1, 'x']");
}

#[test]
fn map_lambda_double() {
    assert_eq!(
        run_print("list(map(lambda x: x * 2, [1, 2, 3]))"),
        "[2, 4, 6]"
    );
}

#[test]
fn list_copy_independent_mutation() {
    assert_eq!(
        run_python_one("a = [1]\nb = a.copy()\nb.append(2)\nprint(a, b)\n"),
        "[1] [1, 2]"
    );
}

#[test]
fn truthy_filter_in_comprehension() {
    assert_eq!(
        run_print("[x for x in [0, 1, 2, '', 'a'] if x]"),
        "[1, 2, 'a']"
    );
}

#[test]
fn membership_test_in_comprehension() {
    assert_eq!(
        run_print("[c for c in 'hello' if c in 'aeiou']"),
        "['e', 'o']"
    );
}

#[test]
fn pass_in_empty_function() {
    assert_eq!(
        run_python_one("def f():\n pass\nprint(callable(f))\n"),
        "True"
    );
}

#[test]
fn ellipsis_singleton() {
    assert_eq!(run_print("... is ..."), "True");
}

#[test]
fn next_on_iterator_with_default() {
    assert_eq!(
        run_python_one("it = iter([1])\nprint(next(it, 99))\nprint(next(it, 99))\n"),
        "1\n99"
    );
}

#[test]
fn reversed_iteration() {
    assert_eq!(run_print("list(reversed([1, 2, 3]))"), "[3, 2, 1]");
}

#[test]
fn join_with_genexp() {
    assert_eq!(
        run_python_one("print('-'.join(str(x) for x in range(3)))\n"),
        "0-1-2"
    );
}

#[test]
fn fstring_with_conditional_expression() {
    assert_eq!(
        run_python_one("n = 4\nprint(f'{'even' if n % 2 == 0 else 'odd'}')\n"),
        "even"
    );
}

#[test]
fn lambda_in_sorted_key() {
    assert_eq!(
        run_print("sorted([(2, 'b'), (1, 'a')], key=lambda t: t[1])"),
        "[(1, 'a'), (2, 'b')]"
    );
}

#[test]
fn dict_get_with_default() {
    assert_eq!(run_print("{'a': 1}.get('b', 0)"), "0");
}

#[test]
fn set_intersection_update_operator() {
    assert_eq!(
        run_python_one("s = {1, 2, 3}\ns &= {2, 3, 4}\nprint(sorted(s))\n"),
        "[2, 3]"
    );
}
