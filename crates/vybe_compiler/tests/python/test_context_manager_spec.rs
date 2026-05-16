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

compile_case!(ctx_enter_exit_compile, "class Ctx:\n    def __enter__(self):\n        return self\n    def __exit__(self, exc_type, exc, tb):\n        return False\n");
compile_case!(ctx_enter_value_compile, "class Ctx:\n    def __enter__(self):\n        return 42\n    def __exit__(self, exc_type, exc, tb):\n        return False\nwith Ctx() as value:\n    print(value)\n");
compile_case!(ctx_multiple_nested_compile, "with open('a') as f:\n    with open('b') as g:\n        with open('c') as h:\n            pass\n");
compile_case!(ctx_multiple_same_line_compile, "with open('a') as f, open('b') as g, open('c') as h:\n    pass\n");
compile_case!(ctx_suppress_exception_compile, "class Ctx:\n    def __enter__(self):\n        return self\n    def __exit__(self, exc_type, exc, tb):\n        return True\nwith Ctx():\n    raise ValueError('boom')\n");
compile_case!(ctx_no_as_compile, "class Lock:\n    def __enter__(self):\n        return self\n    def __exit__(self, exc_type, exc, tb):\n        return False\nwith Lock():\n    pass\n");
compile_case!(ctx_tuple_target_compile, "class Ctx:\n    def __enter__(self):\n        return (1, 2)\n    def __exit__(self, exc_type, exc, tb):\n        return False\nwith Ctx() as (a, b):\n    print(a, b)\n");
compile_case!(ctx_return_none_compile, "class Ctx:\n    def __enter__(self):\n        return None\n    def __exit__(self, exc_type, exc, tb):\n        return False\nwith Ctx() as value:\n    pass\n");
compile_case!(ctx_exception_args_compile, "class Ctx:\n    def __enter__(self):\n        return self\n    def __exit__(self, exc_type, exc, tb):\n        print(exc_type, exc, tb)\n        return False\n");
compile_case!(ctx_rethrow_compile, "class Ctx:\n    def __enter__(self):\n        return self\n    def __exit__(self, exc_type, exc, tb):\n        return False\nwith Ctx():\n    risky()\n");
compile_case!(async_ctx_enter_exit_compile, "class ACtx:\n    async def __aenter__(self):\n        return self\n    async def __aexit__(self, exc_type, exc, tb):\n        return False\n");
compile_case!(async_ctx_basic_compile, "async def main():\n    async with resource() as r:\n        print(r)\n");
compile_case!(async_ctx_nested_compile, "async def main():\n    async with first() as a:\n        async with second() as b:\n            print(a, b)\n");
compile_case!(async_ctx_multiple_compile, "async def main():\n    async with first() as a, second() as b:\n        print(a, b)\n");
compile_case!(ctx_in_try_compile, "try:\n    with open('x') as f:\n        pass\nfinally:\n    cleanup()\n");
compile_case!(ctx_in_loop_compile, "for name in ['a', 'b']:\n    with open(name) as f:\n        print(name)\n");
compile_case!(ctx_in_function_compile, "def read(name):\n    with open(name) as f:\n        return f.read()\n");
compile_case!(ctx_manager_expr_compile, "factory = open\nwith factory('x') as f:\n    pass\n");
compile_case!(ctx_attr_manager_compile, "obj = resource_holder\nwith obj.manager() as r:\n    pass\n");
compile_case!(ctx_line_break_compile, "with open('a') as f,\n     open('b') as g:\n    pass\n");
compile_case!(ctx_line_break_paren_compile, "with (\n    open('a') as f,\n    open('b') as g\n):\n    pass\n");
compile_case!(ctx_generator_style_compile, "from contextlib import contextmanager\n@contextmanager\ndef cm():\n    yield 1\n");
compile_case!(ctx_async_generator_style_compile, "from contextlib import asynccontextmanager\n@asynccontextmanager\nasync def acm():\n    yield 1\n");
runtime_case!(with_none_target_runtime, "class Dummy:\n    pass\nvalue = None\nif value is None:\n    print('none')\n", "none");
runtime_case!(with_fallback_runtime, "x = 0\nif not x:\n    print('fallback')\n", "fallback");
compile_case!(ctx_exit_true_compile, "class Ctx:\n    def __enter__(self):\n        return self\n    def __exit__(self, exc_type, exc, tb):\n        return True\n");
compile_case!(ctx_async_exit_true_compile, "class ACtx:\n    async def __aenter__(self):\n        return self\n    async def __aexit__(self, exc_type, exc, tb):\n        return True\n");
compile_case!(ctx_enter_returns_self_compile, "class Ctx:\n    def __enter__(self):\n        return self\n    def __exit__(self, exc_type, exc, tb):\n        return False\n");
compile_case!(ctx_enter_returns_tuple_compile, "class Ctx:\n    def __enter__(self):\n        return (1, 2, 3)\n    def __exit__(self, exc_type, exc, tb):\n        return False\n");
compile_case!(ctx_async_with_await_compile, "async def main():\n    async with await make_ctx() as ctx:\n        print(ctx)\n");