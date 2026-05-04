use super::helpers::*;

macro_rules! runtime_case {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_python_one($src), $expected);
        }
    };
}

macro_rules! compile_case {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

runtime_case!(iter_list_next_runtime, "it = iter([1, 2, 3])\nprint(next(it))\n", "1");
runtime_case!(iter_string_next_runtime, "it = iter('abc')\nprint(next(it))\n", "a");
runtime_case!(iter_tuple_next_runtime, "it = iter((1, 2))\nprint(next(it))\n", "1");
runtime_case!(iter_range_next_runtime, "it = iter(range(5))\nprint(next(it))\n", "0");
compile_case!(iter_default_compile, "it = iter([1, 2, 3])\nx = next(it, None)\n");
compile_case!(iter_custom_iter_compile, "class Counter:\n    def __iter__(self):\n        return self\n    def __next__(self):\n        raise StopIteration\n");
compile_case!(iter_custom_for_compile, "class Counter:\n    def __iter__(self):\n        return self\n    def __next__(self):\n        raise StopIteration\nfor item in Counter():\n    print(item)\n");
compile_case!(generator_send_compile, "def gen():\n    value = yield 1\n    yield value\n\ng = gen()\ng.send(None)\n");
compile_case!(generator_send_value_compile, "def gen():\n    x = yield 1\n    yield x\n\ng = gen()\nnext(g)\ng.send(42)\n");
compile_case!(generator_throw_compile, "def gen():\n    try:\n        yield 1\n    except ValueError:\n        yield 2\n\ng = gen()\nnext(g)\ng.throw(ValueError())\n");
compile_case!(generator_close_compile, "def gen():\n    yield 1\n\ng = gen()\ng.close()\n");
compile_case!(generator_return_value_compile, "def gen():\n    yield 1\n    return 99\n");
compile_case!(generator_try_finally_compile, "def gen():\n    try:\n        yield 1\n    finally:\n        cleanup()\n");
compile_case!(generator_expression_nested_compile, "x = ((i, j) for i in range(2) for j in range(2))\n");
compile_case!(generator_expression_filter_compile, "x = (i for i in range(10) if i % 2 == 0)\n");
compile_case!(generator_expression_ifelse_compile, "x = ('even' if i % 2 == 0 else 'odd' for i in range(4))\n");
compile_case!(yield_from_subgenerator_compile, "def sub():\n    yield 1\n    yield 2\ndef gen():\n    yield from sub()\n");
compile_case!(yield_from_list_compile, "def gen():\n    yield from [1, 2, 3]\n");
compile_case!(yield_in_try_except_compile, "def gen():\n    try:\n        yield 1\n    except Exception:\n        yield 2\n");
compile_case!(yield_in_with_compile, "def gen():\n    with open('x') as f:\n        yield f.read()\n");
compile_case!(yield_in_comprehension_compile, "def gen():\n    for x in [1, 2, 3]:\n        yield x * 2\n");
runtime_case!(enumerate_iterator_runtime, "it = enumerate(['a', 'b'])\nprint(next(it)[0])\n", "0");
runtime_case!(zip_iterator_runtime, "it = zip([1, 2], ['a', 'b'])\nprint(next(it)[1])\n", "a");
compile_case!(reversed_iterator_compile, "it = reversed([1, 2, 3])\n");
compile_case!(iter_callable_sentinel_compile, "def reader():\n    return 0\nit = iter(reader, 0)\n");
compile_case!(async_generator_basic_compile, "async def agen():\n    yield 1\n");
compile_case!(async_generator_async_for_compile, "async def main():\n    async for value in agen():\n        print(value)\n");
compile_case!(async_generator_yield_from_like_compile, "async def agen():\n    for x in [1, 2]:\n        yield x\n");
compile_case!(stop_iteration_custom_compile, "class Done(Exception):\n    pass\nclass I:\n    def __iter__(self):\n        return self\n    def __next__(self):\n        raise StopIteration\n");
compile_case!(iter_over_dict_keys_compile, "for key in {'a': 1, 'b': 2}:\n    print(key)\n");