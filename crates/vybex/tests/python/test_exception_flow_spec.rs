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

compile_case!(try_else_finally_compile, "try:\n    risky()\nexcept ValueError:\n    handle()\nelse:\n    cleanup()\nfinally:\n    finish()\n");
compile_case!(try_multiple_except_else_compile, "try:\n    risky()\nexcept ValueError:\n    pass\nexcept TypeError:\n    pass\nelse:\n    done()\n");
compile_case!(raise_custom_exception_compile, "class MyError(Exception):\n    pass\nraise MyError('boom')\n");
compile_case!(raise_exception_args_compile, "raise ValueError('bad', 42)\n");
compile_case!(except_baseexception_compile, "try:\n    risky()\nexcept BaseException:\n    pass\n");
compile_case!(except_name_shadow_compile, "try:\n    risky()\nexcept ValueError as value_error:\n    print(value_error)\n");
compile_case!(finally_return_compile, "def f():\n    try:\n        return 1\n    finally:\n        return 2\n");
compile_case!(finally_break_compile, "for i in range(3):\n    try:\n        pass\n    finally:\n        break\n");
compile_case!(finally_continue_compile, "for i in range(3):\n    try:\n        pass\n    finally:\n        continue\n");
compile_case!(nested_finally_compile, "try:\n    try:\n        risky()\n    finally:\n        clean_inner()\nfinally:\n    clean_outer()\n");
compile_case!(assert_raises_custom_compile, "class MyError(Exception):\n    pass\ntry:\n    raise MyError()\nexcept MyError:\n    pass\n");
compile_case!(exception_context_compile, "try:\n    raise ValueError()\nexcept ValueError as exc:\n    ctx = exc.__context__\n");
compile_case!(exception_cause_compile, "try:\n    raise ValueError()\nexcept ValueError as exc:\n    raise RuntimeError() from exc\n");
compile_case!(exception_args_compile, "exc = ValueError('bad', 3)\nargs = exc.args\n");
compile_case!(except_star_compile, "try:\n    raise ExceptionGroup('group', [ValueError('a')])\nexcept* ValueError:\n    pass\n");
compile_case!(exception_group_compile, "group = ExceptionGroup('many', [ValueError('a'), TypeError('b')])\n");
compile_case!(raise_note_compile, "exc = ValueError('bad')\nexc.add_note('context')\nraise exc\n");
runtime_case!(try_else_runtime, "try:\n    x = 1\nexcept Exception:\n    print('except')\nelse:\n    print('else')\n", "else");
runtime_case!(finally_overrides_return_runtime, "def f():\n    try:\n        return 1\n    finally:\n        return 2\nprint(f())\n", "2");
runtime_case!(assert_runtime_true, "assert 1 == 1\nprint('ok')\n", "ok");
runtime_case!(raise_caught_runtime, "try:\n    raise ValueError('bad')\nexcept ValueError:\n    print('caught')\n", "caught");
runtime_case!(nested_except_runtime, "try:\n    try:\n        raise ValueError()\n    except ValueError:\n        print('inner')\nexcept Exception:\n    print('outer')\n", "inner");
compile_case!(return_in_except_compile, "def f():\n    try:\n        risky()\n    except Exception:\n        return 1\n");
compile_case!(raise_in_finally_compile, "try:\n    risky()\nfinally:\n    raise RuntimeError('final')\n");
compile_case!(except_tuple_custom_compile, "class A(Exception): pass\nclass B(Exception): pass\ntry:\n    risky()\nexcept (A, B):\n    pass\n");
compile_case!(try_except_in_comprehension_compile, "def f(xs):\n    out = []\n    for x in xs:\n        try:\n            out.append(x)\n        except Exception:\n            pass\n    return out\n");
compile_case!(contextlib_suppress_compile, "from contextlib import suppress\nwith suppress(ValueError):\n    raise ValueError()\n");
compile_case!(exitstack_compile, "from contextlib import ExitStack\nwith ExitStack() as stack:\n    pass\n");
compile_case!(reraise_bare_compile, "try:\n    risky()\nexcept Exception:\n    raise\n");
compile_case!(reraise_named_compile, "try:\n    risky()\nexcept Exception as exc:\n    raise exc\n");