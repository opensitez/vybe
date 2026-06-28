use crate::helpers::{run_print, run_python_one};

#[test]
fn lambda_identity() {
    assert_eq!(run_python_one("f = lambda x: x\nprint(f(7))\n"), "7");
}

#[test]
fn lambda_add_two_args() {
    assert_eq!(
        run_python_one("f = lambda a, b: a + b\nprint(f(2, 5))\n"),
        "7"
    );
}

#[test]
fn lambda_multiply() {
    assert_eq!(run_python_one("f = lambda x: x * 3\nprint(f(4))\n"), "12");
}

#[test]
fn lambda_no_args() {
    assert_eq!(run_python_one("f = lambda: 99\nprint(f())\n"), "99");
}

#[test]
fn lambda_default_arg() {
    assert_eq!(
        run_python_one("f = lambda x, y=10: x + y\nprint(f(5))\n"),
        "15"
    );
}

#[test]
fn lambda_in_list_comp() {
    assert_eq!(
        run_print("[(lambda x: x + 1)(i) for i in range(3)]"),
        "[1, 2, 3]"
    );
}

#[test]
fn lambda_sorted_key_length() {
    assert_eq!(
        run_print("sorted(['bb', 'a', 'ccc'], key=lambda s: len(s))"),
        "['a', 'bb', 'ccc']"
    );
}

#[test]
fn lambda_sorted_key_second_element() {
    assert_eq!(
        run_print("sorted([(2, 'b'), (1, 'a')], key=lambda t: t[1])"),
        "[(1, 'a'), (2, 'b')]"
    );
}

#[test]
fn lambda_map_square() {
    assert_eq!(
        run_print("list(map(lambda x: x * x, [1, 2, 3]))"),
        "[1, 4, 9]"
    );
}

#[test]
fn lambda_filter_positive() {
    assert_eq!(
        run_print("list(filter(lambda x: x > 0, [-1, 0, 2]))"),
        "[2]"
    );
}

#[test]
fn lambda_filter_truthy_strings() {
    assert_eq!(
        run_print("list(filter(lambda s: bool(s), ['', 'a', '']))"),
        "['a']"
    );
}

#[test]
fn lambda_closure_captures_outer() {
    assert_eq!(
        run_python_one("def mk(n):\n return lambda x: x + n\nprint(mk(3)(4))\n"),
        "7"
    );
}

#[test]
fn lambda_closure_captures_loop_var_default() {
    assert_eq!(
        run_python_one("funcs = [lambda x=i: x for i in range(3)]\nprint(funcs[2]())\n"),
        "2"
    );
}

#[test]
fn lambda_returned_from_function() {
    assert_eq!(
        run_python_one("def twice():\n return lambda x: x * 2\nprint(twice()(5))\n"),
        "10"
    );
}

#[test]
fn lambda_conditional_expression() {
    assert_eq!(
        run_python_one("f = lambda x: 'big' if x > 5 else 'small'\nprint(f(10), f(1))\n"),
        "big small"
    );
}

#[test]
fn lambda_bool_logic() {
    assert_eq!(
        run_python_one("f = lambda a, b: a and b\nprint(f(True, False))\n"),
        "False"
    );
}

#[test]
fn lambda_string_upper() {
    assert_eq!(
        run_python_one("f = lambda s: s.upper()\nprint(f('ab'))\n"),
        "AB"
    );
}

#[test]
fn lambda_list_append_side_effect() {
    assert_eq!(
        run_python_one("acc = []\nf = lambda x: acc.append(x)\nf(1)\nf(2)\nprint(acc)\n"),
        "[1, 2]"
    );
}

#[test]
fn lambda_as_dict_value() {
    assert_eq!(
        run_python_one("d = {'inc': lambda x: x + 1}\nprint(d['inc'](4))\n"),
        "5"
    );
}

#[test]
fn lambda_nested_call() {
    assert_eq!(
        run_python_one("f = lambda g, x: g(x)\nprint(f(lambda y: y + 1, 3))\n"),
        "4"
    );
}

#[test]
fn lambda_tuple_return() {
    assert_eq!(run_python_one("f = lambda: (1, 2)\nprint(f())\n"), "(1, 2)");
}

#[test]
fn lambda_max_key() {
    assert_eq!(
        run_python_one("print(max(['a', 'bbb'], key=lambda s: len(s)))\n"),
        "bbb"
    );
}

#[test]
fn lambda_min_key() {
    assert_eq!(
        run_python_one("print(min([3, 1, 2], key=lambda x: -x))\n"),
        "3"
    );
}

#[test]
fn lambda_any_predicate() {
    assert_eq!(
        run_python_one("print(any(map(lambda x: x > 2, [1, 2, 3])))\n"),
        "True"
    );
}

#[test]
fn lambda_all_predicate() {
    assert_eq!(
        run_python_one("print(all(map(lambda x: x > 0, [1, 2])))\n"),
        "True"
    );
}

#[test]
fn lambda_reduce_style_manual() {
    assert_eq!(
        run_python_one(
            "acc = 0\nfor v in [1, 2, 3]:\n acc = (lambda a, b: a + b)(acc, v)\nprint(acc)\n"
        ),
        "6"
    );
}

#[test]
fn lambda_unpack_args() {
    assert_eq!(
        run_python_one("f = lambda a, b: a - b\nprint(f(*[5, 2]))\n"),
        "3"
    );
}

#[test]
fn lambda_with_none_return() {
    assert_eq!(run_python_one("f = lambda: None\nprint(f())\n"), "None");
}

#[test]
fn lambda_compare_chain() {
    assert_eq!(
        run_python_one("f = lambda x: 0 < x < 10\nprint(f(5), f(11))\n"),
        "True False"
    );
}

#[test]
fn lambda_modulo_predicate() {
    assert_eq!(
        run_print("list(filter(lambda x: x % 2 == 0, range(6)))"),
        "[0, 2, 4]"
    );
}

#[test]
fn lambda_string_startswith() {
    assert_eq!(
        run_print("list(filter(lambda s: s.startswith('a'), ['ab', 'ba']))"),
        "['ab']"
    );
}

#[test]
fn lambda_index_in_expression() {
    assert_eq!(
        run_python_one("f = lambda xs: xs[0]\nprint(f([9, 8]))\n"),
        "9"
    );
}

#[test]
fn lambda_slice_in_expression() {
    assert_eq!(
        run_python_one("f = lambda xs: xs[:2]\nprint(f([1, 2, 3]))\n"),
        "[1, 2]"
    );
}

#[test]
fn lambda_dict_get_default() {
    assert_eq!(
        run_python_one("f = lambda d, k: d.get(k, 0)\nprint(f({'a': 1}, 'b'))\n"),
        "0"
    );
}

#[test]
fn lambda_power() {
    assert_eq!(
        run_python_one("f = lambda x, n: x ** n\nprint(f(2, 4))\n"),
        "16"
    );
}

#[test]
fn lambda_floor_div() {
    assert_eq!(
        run_python_one("f = lambda a, b: a // b\nprint(f(7, 2))\n"),
        "3"
    );
}

#[test]
fn lambda_abs_via_conditional() {
    assert_eq!(
        run_python_one("f = lambda x: -x if x < 0 else x\nprint(f(-3))\n"),
        "3"
    );
}

#[test]
fn lambda_join_strings() {
    assert_eq!(
        run_python_one("f = lambda parts: '-'.join(parts)\nprint(f(['a', 'b']))\n"),
        "a-b"
    );
}

#[test]
fn lambda_enumerate_build() {
    assert_eq!(
        run_python_one(
            "pairs = list(map(lambda t: t[0] + t[1], enumerate(['x', 'y'])))\nprint(pairs)\n"
        ),
        "[0, 1]"
    );
}

#[test]
fn lambda_zip_sum_pairs() {
    assert_eq!(
        run_python_one("f = lambda a, b: a + b\nprint(list(map(f, [1, 2], [10, 20])))\n"),
        "[11, 22]"
    );
}

#[test]
fn lambda_in_sorted_reverse_key() {
    assert_eq!(
        run_print("sorted([1, 3, 2], key=lambda x: x, reverse=True)"),
        "[3, 2, 1]"
    );
}

#[test]
fn lambda_call_twice_stateless() {
    assert_eq!(
        run_python_one("f = lambda x: x + 1\nprint(f(1), f(1))\n"),
        "2 2"
    );
}

#[test]
fn lambda_higher_order_returns_lambda() {
    assert_eq!(
        run_python_one(
            "def compose(f, g):\n return lambda x: f(g(x))\nprint(compose(lambda x: x+1, lambda x: x*2)(3))\n"
        ),
        "7"
    );
}

#[test]
fn lambda_with_walrus_inside() {
    assert_eq!(
        run_python_one("f = lambda x: (y := x * 2)\nprint(f(3))\n"),
        "6"
    );
}

#[test]
fn lambda_exception_in_body_uncaught() {
    assert_eq!(
        run_python_one("try:\n (lambda: 1/0)()\nexcept ZeroDivisionError:\n print('lam')\n"),
        "lam"
    );
}
