use crate::helpers::{run_python_one, run_print};

#[test]
fn generator_yield_single() {
    assert_eq!(
        run_python_one("def g():\n yield 1\nprint(list(g()))\n"),
        "[1]"
    );
}

#[test]
fn generator_yield_multiple() {
    assert_eq!(
        run_python_one("def g():\n yield 1\n yield 2\nprint(list(g()))\n"),
        "[1, 2]"
    );
}

#[test]
fn generator_yield_from_loop() {
    assert_eq!(
        run_python_one("def g():\n for i in range(3):\n  yield i\nprint(list(g()))\n"),
        "[0, 1, 2]"
    );
}

#[test]
fn generator_next_manual() {
    assert_eq!(
        run_python_one("def g():\n yield 10\nit = g()\nprint(next(it))\n"),
        "10"
    );
}

#[test]
fn generator_stop_iteration() {
    assert_eq!(
        run_python_one("def g():\n return\n yield 1\ntry:\n next(g())\nexcept StopIteration:\n print('stop')\n"),
        "stop"
    );
}

#[test]
fn generator_send_not_required_basic() {
    assert_eq!(
        run_python_one("def g():\n yield 5\nprint(next(g()))\n"),
        "5"
    );
}

#[test]
fn generator_yield_from_subgenerator() {
    assert_eq!(
        run_python_one("def inner():\n yield 2\ndef outer():\n yield 1\n yield from inner()\n yield 3\nprint(list(outer()))\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn generator_expression_in_sum() {
    assert_eq!(
        run_python_one("print(sum(x * x for x in range(4)))\n"),
        "14"
    );
}

#[test]
fn generator_expression_in_max() {
    assert_eq!(
        run_python_one("print(max(len(s) for s in ['a', 'bbb']))\n"),
        "3"
    );
}

#[test]
fn generator_expression_in_any() {
    assert_eq!(
        run_python_one("print(any(x > 2 for x in [1, 2, 3]))\n"),
        "True"
    );
}

#[test]
fn generator_expression_in_all() {
    assert_eq!(
        run_python_one("print(all(x > 0 for x in [1, 2]))\n"),
        "True"
    );
}

#[test]
fn generator_expression_filtered() {
    assert_eq!(
        run_print("list(x for x in range(5) if x % 2 == 0)"),
        "[0, 2, 4]"
    );
}

#[test]
fn generator_lazy_not_materialized_until_iter() {
    assert_eq!(
        run_python_one("def g():\n yield 1\ngen = (x for x in g())\nprint(next(gen))\n"),
        "1"
    );
}

#[test]
fn generator_list_materialize() {
    assert_eq!(
        run_print("list(x + 1 for x in range(3))"),
        "[1, 2, 3]"
    );
}

#[test]
fn generator_tuple_materialize() {
    assert_eq!(
        run_print("tuple(x for x in range(2))"),
        "(0, 1)"
    );
}

#[test]
fn generator_set_materialize() {
    assert_eq!(
        run_print("sorted({x for x in [3, 1, 2, 1]})"),
        "[1, 2, 3]"
    );
}

#[test]
fn generator_dict_companion_genexp_keys() {
    assert_eq!(
        run_print("{k: k for k in (x for x in range(2))}"),
        "{0: 0, 1: 1}"
    );
}

#[test]
fn generator_yield_none_explicit() {
    assert_eq!(
        run_python_one("def g():\n yield None\nprint(list(g()))\n"),
        "[None]"
    );
}

#[test]
fn generator_return_value_captured_in_stopiteration() {
    assert_eq!(
        run_python_one("def g():\n yield 1\n return 99\nit = g()\nprint(next(it))\ntry:\n next(it)\nexcept StopIteration as e:\n print(e.value)\n"),
        "1\n99"
    );
}

#[test]
fn generator_nested_yield() {
    assert_eq!(
        run_python_one("def g():\n def inner():\n  yield 2\n yield 1\n yield from inner()\nprint(list(g()))\n"),
        "[1, 2]"
    );
}

#[test]
fn generator_in_for_loop() {
    assert_eq!(
        run_python_one("def g():\n yield 'a'\n yield 'b'\nout = ''\nfor ch in g():\n out += ch\nprint(out)\n"),
        "ab"
    );
}

#[test]
fn generator_break_stops_iteration() {
    assert_eq!(
        run_python_one("def g():\n for i in range(5):\n  yield i\ncount = 0\nfor _ in g():\n count += 1\n if count == 2:\n  break\nprint(count)\n"),
        "2"
    );
}

#[test]
fn generator_enumerate_over_gen() {
    assert_eq!(
        run_python_one("def g():\n yield 10\n yield 20\nprint(list(enumerate(g())))\n"),
        "[(0, 10), (1, 20)]"
    );
}

#[test]
fn generator_zip_with_gen() {
    assert_eq!(
        run_python_one("def g():\n yield 1\n yield 2\nprint(list(zip(g(), ['a', 'b'])))\n"),
        "[(1, 'a'), (2, 'b')]"
    );
}

#[test]
fn generator_chain_two() {
    assert_eq!(
        run_python_one("def a():\n yield 1\ndef b():\n yield 2\nprint(list(a()) + list(b()))\n"),
        "[1, 2]"
    );
}

#[test]
fn generator_filter_on_genexp() {
    assert_eq!(
        run_print("list(filter(lambda x: x > 1, (i for i in range(4))))"),
        "[2, 3]"
    );
}

#[test]
fn generator_map_on_genexp() {
    assert_eq!(
        run_print("list(map(lambda x: x * 2, (i for i in range(3))))"),
        "[0, 2, 4]"
    );
}

#[test]
fn generator_fibonacci_style() {
    assert_eq!(
        run_python_one("def fib():\n a, b = 0, 1\n while a < 10:\n  yield a\n  a, b = b, a + b\nprint(list(fib()))\n"),
        "[0, 1, 1, 2, 3, 5, 8]"
    );
}

#[test]
fn generator_count_with_sentinel() {
    assert_eq!(
        run_python_one("def count(n):\n while n > 0:\n  yield n\n  n -= 1\nprint(list(count(3)))\n"),
        "[3, 2, 1]"
    );
}

#[test]
fn generator_read_file_lines_style() {
    assert_eq!(
        run_python_one("def lines():\n for s in ['a', 'b']:\n  yield s.upper()\nprint(list(lines()))\n"),
        "['A', 'B']"
    );
}

#[test]
fn generator_yield_from_list() {
    assert_eq!(
        run_python_one("def g():\n yield from [1, 2, 3]\nprint(list(g()))\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn generator_yield_from_range() {
    assert_eq!(
        run_python_one("def g():\n yield from range(3)\nprint(list(g()))\n"),
        "[0, 1, 2]"
    );
}

#[test]
fn generator_close_raises_generator_exit() {
    assert_eq!(
        run_python_one("def g():\n try:\n  yield 1\n finally:\n  print('fin')\nit = g()\nprint(next(it))\nit.close()\n"),
        "1\nfin"
    );
}

#[test]
fn generator_throw_into_generator() {
    assert_eq!(
        run_python_one("def g():\n try:\n  yield 1\n except ValueError:\n  yield 'recovered'\nit = g()\nprint(next(it))\nprint(it.throw(ValueError))\n"),
        "1\nrecovered"
    );
}

#[test]
fn generator_state_persists_between_yields() {
    assert_eq!(
        run_python_one("def g():\n x = 0\n while x < 3:\n  yield x\n  x += 1\nprint(list(g()))\n"),
        "[0, 1, 2]"
    );
}

#[test]
fn generator_multiple_iterators_independent() {
    assert_eq!(
        run_python_one("def g():\n yield 1\na = g()\nb = g()\nprint(next(a), next(b))\n"),
        "1 1"
    );
}

#[test]
fn generator_same_iterator_exhausted() {
    assert_eq!(
        run_python_one("def g():\n yield 1\nit = g()\nprint(list(it), list(it))\n"),
        "[1] []"
    );
}

#[test]
fn generator_comprehension_scope_local() {
    assert_eq!(
        run_python_one("out = (x for x in range(2))\nprint(list(out))\n"),
        "[0, 1]"
    );
}

#[test]
fn generator_with_condition_on_length() {
    assert_eq!(
        run_python_one("print(list(s for s in ['a', 'bb'] if len(s) == 1))\n"),
        "['a']"
    );
}

#[test]
fn generator_join_strings() {
    assert_eq!(
        run_python_one("print('-'.join(str(x) for x in range(3)))\n"),
        "0-1-2"
    );
}

#[test]
fn generator_min_of_genexp() {
    assert_eq!(
        run_python_one("print(min(x for x in [3, 1, 2]))\n"),
        "1"
    );
}

#[test]
fn generator_sorted_materialize() {
    assert_eq!(
        run_python_one("print(sorted(x for x in [3, 1, 2]))\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn generator_bool_on_gen_object() {
    assert_eq!(
        run_python_one("g = (x for x in range(1))\nprint(bool(g))\n"),
        "True"
    );
}

#[test]
fn generator_len_not_supported() {
    assert_eq!(
        run_python_one("g = (x for x in range(3))\ntry:\n len(g)\nexcept TypeError:\n print('no')\n"),
        "no"
    );
}

#[test]
fn generator_reversed_on_list_not_gen() {
    assert_eq!(
        run_python_one("print(list(reversed([1, 2, 3])))\n"),
        "[3, 2, 1]"
    );
}
