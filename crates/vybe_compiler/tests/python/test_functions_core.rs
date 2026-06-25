use crate::helpers::{run_python_one, run_print};

#[test]
fn function_return_literal() {
    assert_eq!(
        run_python_one("def f():\n return 42\nprint(f())\n"),
        "42"
    );
}

#[test]
fn function_return_expression() {
    assert_eq!(
        run_python_one("def add(a, b):\n return a + b\nprint(add(3, 4))\n"),
        "7"
    );
}

#[test]
fn function_default_arg_used() {
    assert_eq!(
        run_python_one("def g(n=10):\n return n\nprint(g())\n"),
        "10"
    );
}

#[test]
fn function_default_arg_overridden() {
    assert_eq!(
        run_python_one("def g(n=10):\n return n\nprint(g(5))\n"),
        "5"
    );
}

#[test]
fn function_multiple_defaults() {
    assert_eq!(
        run_python_one("def f(a, b=2, c=3):\n return a + b + c\nprint(f(1))\n"),
        "6"
    );
}

#[test]
fn function_early_return() {
    assert_eq!(
        run_python_one("def f(x):\n if x < 0:\n  return 0\n return x\nprint(f(-1))\n"),
        "0"
    );
}

#[test]
fn function_no_explicit_return_gives_none() {
    assert_eq!(
        run_python_one("def f():\n pass\nprint(f())\n"),
        "None"
    );
}

#[test]
fn function_local_shadows_parameter() {
    assert_eq!(
        run_python_one("def f(x):\n x = x + 1\n return x\nprint(f(1))\n"),
        "2"
    );
}

#[test]
fn nested_function_closure() {
    assert_eq!(
        run_python_one("def outer(x):\n def inner():\n  return x\n return inner()\nprint(outer(9))\n"),
        "9"
    );
}

#[test]
fn nested_function_calls_outer_param() {
    assert_eq!(
        run_python_one("def outer(a, b):\n def inner():\n  return a * b\n return inner()\nprint(outer(3, 4))\n"),
        "12"
    );
}

#[test]
fn function_recursion_factorial() {
    assert_eq!(
        run_python_one("def fact(n):\n if n <= 1:\n  return 1\n return n * fact(n - 1)\nprint(fact(5))\n"),
        "120"
    );
}

#[test]
fn function_recursion_fibonacci() {
    assert_eq!(
        run_python_one("def fib(n):\n if n < 2:\n  return n\n return fib(n-1) + fib(n-2)\nprint(fib(6))\n"),
        "8"
    );
}

#[test]
fn function_varargs_sum() {
    assert_eq!(
        run_python_one("def total(*args):\n return sum(args)\nprint(total(1, 2, 3))\n"),
        "6"
    );
}

#[test]
fn function_kwargs_lookup() {
    assert_eq!(
        run_python_one("def f(**kw):\n return kw['a']\nprint(f(a=7))\n"),
        "7"
    );
}

#[test]
fn function_mixed_positional_and_kwargs() {
    assert_eq!(
        run_python_one("def f(x, y=0):\n return x + y\nprint(f(2, y=3))\n"),
        "5"
    );
}

#[test]
fn function_returns_list_mutation_visible() {
    assert_eq!(
        run_print("def f():\n return [1]\nx = f()\nx.append(2)\nx"),
        "[1, 2]"
    );
}

#[test]
fn function_returns_new_list_each_call() {
    assert_eq!(
        run_python_one("def f():\n return []\na = f()\nb = f()\nprint(a is b)\n"),
        "False"
    );
}

#[test]
fn function_call_in_expression() {
    assert_eq!(
        run_python_one("def dbl(x):\n return x * 2\nprint(1 + dbl(2))\n"),
        "5"
    );
}

#[test]
fn function_as_callback_map() {
    assert_eq!(
        run_print("def inc(x):\n return x + 1\nlist(map(inc, [1, 2, 3]))"),
        "[2, 3, 4]"
    );
}

#[test]
fn function_as_predicate_filter() {
    assert_eq!(
        run_print("def is_pos(x):\n return x > 0\nlist(filter(is_pos, [-1, 0, 2]))"),
        "[2]"
    );
}

#[test]
fn function_mutual_recursion_even() {
    assert_eq!(
        run_python_one(
            "def is_even(n):\n if n == 0:\n  return True\n return is_odd(n - 1)\n\
             def is_odd(n):\n if n == 0:\n  return False\n return is_even(n - 1)\nprint(is_even(4))\n"
        ),
        "True"
    );
}

#[test]
fn function_docstring_does_not_affect_return() {
    assert_eq!(
        run_python_one("def f():\n '''docs'''\n return 1\nprint(f())\n"),
        "1"
    );
}

#[test]
fn function_parameter_unpacking() {
    assert_eq!(
        run_python_one("def f(a, b):\n return a + b\nprint(f(*(1, 2)))\n"),
        "3"
    );
}

#[test]
fn function_keyword_unpacking() {
    assert_eq!(
        run_python_one("def f(a, b):\n return a * b\nprint(f(**{'a': 3, 'b': 4}))\n"),
        "12"
    );
}

#[test]
fn function_default_mutable_not_shared_across_calls() {
    assert_eq!(
        run_python_one(
            "def f(x, acc=None):\n if acc is None:\n  acc = []\n acc.append(x)\n return len(acc)\n\
             print(f(1), f(1))\n"
        ),
        "1 1"
    );
}

#[test]
fn function_return_tuple_unpack() {
    assert_eq!(
        run_python_one("def pair():\n return 1, 2\na, b = pair()\nprint(a, b)\n"),
        "1 2"
    );
}

#[test]
fn function_return_conditional_type() {
    assert_eq!(
        run_python_one("def f(flag):\n return 'yes' if flag else 0\nprint(f(True), f(False))\n"),
        "yes 0"
    );
}

#[test]
fn function_local_list_accumulator() {
    assert_eq!(
        run_python_one("def build():\n out = []\n for i in range(3):\n  out.append(i)\n return out\nprint(build())\n"),
        "[0, 1, 2]"
    );
}

#[test]
fn function_pass_by_object_reference() {
    assert_eq!(
        run_python_one("def mutate(xs):\n xs.append(9)\na = [1]\nmutate(a)\nprint(a)\n"),
        "[1, 9]"
    );
}

#[test]
fn function_rebind_parameter_does_not_affect_caller() {
    assert_eq!(
        run_python_one("def f(x):\n x = 99\na = 1\nf(a)\nprint(a)\n"),
        "1"
    );
}

#[test]
fn function_name_in_own_body_recursive() {
    assert_eq!(
        run_python_one("def countdown(n):\n if n <= 0:\n  return 'done'\n return countdown(n - 1)\nprint(countdown(2))\n"),
        "done"
    );
}

#[test]
fn function_with_while_loop_body() {
    assert_eq!(
        run_python_one("def first_gt(threshold):\n n = 0\n while n <= threshold:\n  n += 1\n return n\nprint(first_gt(3))\n"),
        "4"
    );
}

#[test]
fn function_with_for_loop_sum() {
    assert_eq!(
        run_python_one("def sum_range(n):\n total = 0\n for i in range(n):\n  total += i\n return total\nprint(sum_range(4))\n"),
        "6"
    );
}

#[test]
fn function_returns_boolean() {
    assert_eq!(
        run_python_one("def is_empty(xs):\n return len(xs) == 0\nprint(is_empty([]), is_empty([1]))\n"),
        "True False"
    );
}

#[test]
fn function_returns_string_concat() {
    assert_eq!(
        run_python_one("def greet(name):\n return 'hi ' + name\nprint(greet('py'))\n"),
        "hi py"
    );
}

#[test]
fn function_call_chain_nested() {
    assert_eq!(
        run_python_one("def a(x):\n return x + 1\ndef b(x):\n return a(x) * 2\nprint(b(3))\n"),
        "8"
    );
}

#[test]
fn function_optional_none_default() {
    assert_eq!(
        run_python_one("def f(x=None):\n return x is None\nprint(f())\n"),
        "True"
    );
}

#[test]
fn function_bool_return_from_comparison() {
    assert_eq!(
        run_python_one("def eq(a, b):\n return a == b\nprint(eq(1, 1), eq(1, 2))\n"),
        "True False"
    );
}

#[test]
fn function_string_format_in_return() {
    assert_eq!(
        run_python_one("def label(n):\n return f'n={n}'\nprint(label(5))\n"),
        "n=5"
    );
}

#[test]
fn function_multiple_return_paths_same_type() {
    assert_eq!(
        run_python_one("def abs_val(x):\n if x < 0:\n  return -x\n return x\nprint(abs_val(-3), abs_val(2))\n"),
        "3 2"
    );
}

#[test]
fn function_inner_redefines_outer_name_locally() {
    assert_eq!(
        run_python_one("def outer():\n x = 1\n def inner():\n  x = 2\n  return x\n return inner()\nprint(outer())\n"),
        "2"
    );
}

#[test]
fn function_read_outer_without_nonlocal() {
    assert_eq!(
        run_python_one("def outer():\n x = 5\n def inner():\n  return x + 1\n return inner()\nprint(outer())\n"),
        "6"
    );
}

#[test]
fn function_zero_args_print_side_effect() {
    assert_eq!(
        run_python_one("def shout():\n print('ok')\nshout()\n"),
        "ok"
    );
}

#[test]
fn function_returns_dict_literal() {
    assert_eq!(
        run_print("def mapping():\n return {'a': 1}\nmapping()"),
        "{'a': 1}"
    );
}

#[test]
fn function_returns_set_size() {
    assert_eq!(
        run_python_one("def unique(xs):\n return len(set(xs))\nprint(unique([1, 1, 2]))\n"),
        "2"
    );
}
