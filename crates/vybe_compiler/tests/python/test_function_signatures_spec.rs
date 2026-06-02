use super::helpers::*;

macro_rules! compile_case {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

compile_case!(
    positional_only_basic_compile,
    "def add(a, b, /):\n    return a + b\n"
);
compile_case!(
    positional_only_mixed_compile,
    "def add(a, b, /, c):\n    return a + b + c\n"
);
compile_case!(
    positional_only_kwonly_compile,
    "def f(a, /, *, b):\n    return a + b\n"
);
compile_case!(
    bare_star_kwonly_compile,
    "def f(*, flag=False):\n    return flag\n"
);
compile_case!(
    keyword_only_defaults_compile,
    "def f(a, *, left=1, right=2):\n    return a + left + right\n"
);
compile_case!(
    variadic_positional_only_compile,
    "def f(a, /, *args):\n    return args\n"
);
compile_case!(
    variadic_kwonly_compile,
    "def f(*args, sep=':', end='!'):\n    return sep\n"
);
compile_case!(
    param_annotations_defaults_compile,
    "def f(a: int = 1, b: str = 'x') -> str:\n    return b\n"
);
compile_case!(
    return_annotation_union_compile,
    "def f(x: int) -> int | None:\n    return x\n"
);
compile_case!(
    variable_annotation_generic_compile,
    "items: list[dict[str, int]] = []\n"
);
compile_case!(
    annotated_assignment_tuple_compile,
    "coords: tuple[int, int] = (1, 2)\n"
);
compile_case!(
    future_annotations_compile,
    "from __future__ import annotations\ndef f(x: 'Node') -> 'Node':\n    return x\n"
);
compile_case!(type_alias_compile, "Vector = list[float]\n");
compile_case!(type_union_pipe_compile, "value: str | int = 1\n");
compile_case!(async_generator_compile, "async def gen():\n    yield 1\n");
compile_case!(
    async_generator_with_await_compile,
    "async def gen():\n    value = await fetch()\n    yield value\n"
);
compile_case!(
    async_comprehension_if_compile,
    "result = [x async for x in aiter if x > 0]\n"
);
compile_case!(
    with_parenthesized_managers_compile,
    "with (open('a') as f, open('b') as g):\n    pass\n"
);
compile_case!(
    with_line_continuation_compile,
    "with open('a') as f, \\\n+     open('b') as g:\n    pass\n"
);
compile_case!(
    assert_in_function_compile,
    "def check(x):\n    assert x > 0\n"
);
compile_case!(
    assert_message_expr_compile,
    "assert value, 'bad: %s' % value\n"
);
compile_case!(
    assert_in_except_compile,
    "try:\n    risky()\nexcept Exception as exc:\n    assert exc is not None\n"
);
compile_case!(import_dotted_as_compile, "import package.submodule as sm\n");
compile_case!(
    from_import_alias_parenthesized_compile,
    "from package.module import (\n    name as alias,\n    other\n)\n"
);
compile_case!(
    import_relative_alias_compile,
    "from .subpackage import tool as t\n"
);
compile_case!(
    import_future_compile,
    "from __future__ import annotations, generator_stop\n"
);
compile_case!(
    decorator_factory_compile,
    "def deco(name):\n    def wrap(fn):\n        return fn\n    return wrap\n@deco('x')\ndef f():\n    pass\n"
);
compile_case!(
    decorator_with_return_annotation_compile,
    "@cache\ndef f(x: int) -> int:\n    return x\n"
);
compile_case!(lambda_kwonly_compile, "f = lambda x, *, y=1: x + y\n");
compile_case!(
    def_with_annotations_and_unpack_compile,
    "def f(a: int, *args: str, **kwargs: float) -> tuple[int, tuple[str], dict[str, float]]:\n    return a, args, kwargs\n"
);
