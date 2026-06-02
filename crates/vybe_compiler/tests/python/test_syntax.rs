use super::helpers::*;

// ══════════════════════════════════════════════════════════════════════════════
// Walrus operator (:=)
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn walrus_in_if() {
    parse_ok("if (n := 10) > 5:\n    print(n)\n");
}

#[test]
fn walrus_in_while() {
    parse_ok("while chunk := input():\n    process(chunk)\n");
}

#[test]
fn walrus_in_list() {
    parse_ok("results = [y := f(x), y**2, y**3]\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Async / Await
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn async_def() {
    compile_ok("async def fetch():\n    data = await get_data()\n    return data\n");
}

#[test]
fn async_for() {
    compile_ok("async def main():\n    async for item in aiter:\n        print(item)\n");
}

#[test]
fn async_with() {
    compile_ok("async def ctx():\n    async with resource() as r:\n        pass\n");
}

#[test]
fn async_comprehension() {
    parse_ok("result = [x async for x in aiter]\n");
}

#[test]
fn async_dict_comprehension() {
    parse_ok("result = {k: v async for k, v in aitems}\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Decorators
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn decorator_with_args() {
    compile_ok("@app.route('/home')\ndef home():\n    pass\n");
}

#[test]
fn stacked_decorators() {
    compile_ok("@decorator1\n@decorator2\ndef func():\n    pass\n");
}

#[test]
fn decorator_on_class() {
    compile_ok("@dataclass\nclass Point:\n    x: int\n    y: int\n");
}

#[test]
fn decorator_complex_expression() {
    compile_ok("@module.sub.decorator(arg1, key=val)\ndef func():\n    pass\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Type Annotations
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn annotation_function_return() {
    compile_ok("def add(a: int, b: int) -> int:\n    return a + b\n");
}

#[test]
fn annotation_list_type() {
    compile_ok("x: list[int] = [1, 2, 3]\n");
}

#[test]
fn annotation_dict_type() {
    compile_ok("def foo(x: dict[str, list[int]]) -> None:\n    pass\n");
}

#[test]
fn annotation_optional() {
    compile_ok("y: Optional[str] = None\n");
}

#[test]
fn annotation_tuple_type() {
    compile_ok("z: tuple[int, ...] = (1, 2, 3)\n");
}

#[test]
fn annotation_variable_only() {
    compile_ok("x: int\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Global / Nonlocal
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn global_statement() {
    compile_ok("x = 10\ndef change():\n    global x\n    x = 20\n");
}

#[test]
fn nonlocal_statement() {
    compile_ok(
        "def outer():\n    x = 10\n    def inner():\n        nonlocal x\n        x = 20\n    inner()\n",
    );
}

#[test]
fn global_multiple() {
    compile_ok("def f():\n    global a, b, c\n    a = 1\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// With statement details
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn with_multiple_managers() {
    compile_ok("with open('a') as f1, open('b') as f2:\n    pass\n");
}

#[test]
fn with_no_as() {
    compile_ok("with lock:\n    do_stuff()\n");
}

#[test]
fn with_nested() {
    compile_ok("with open('a') as f:\n    with open('b') as g:\n        pass\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Function features
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn args_variadic() {
    compile_ok("def f(*args):\n    print(args)\nf(1, 2, 3)\n");
}

#[test]
fn kwargs_variadic() {
    compile_ok("def f(**kwargs):\n    print(kwargs)\nf(a=1, b=2)\n");
}

#[test]
fn args_and_kwargs() {
    compile_ok("def f(a, b, *args, **kwargs):\n    pass\n");
}

#[test]
fn keyword_only_params() {
    compile_ok("def f(a, *, key=None):\n    pass\n");
}

#[test]
fn call_star_unpack() {
    compile_ok("def f(a, b, c):\n    pass\nargs = [1, 2, 3]\nf(*args)\n");
}

#[test]
fn call_double_star_unpack() {
    compile_ok("def f(a, b):\n    pass\nkw = {'a': 1, 'b': 2}\nf(**kw)\n");
}

#[test]
fn multiple_return_values() {
    let out =
        run_python("def swap(a, b):\n    return b, a\nx, y = swap(1, 2)\nprint(x)\nprint(y)\n");
    assert_eq!(out[0], "2");
    assert_eq!(out[1], "1");
}

#[test]
fn lambda_runtime() {
    assert_eq!(run_python_one("f = lambda x: x + 1\nprint(f(5))\n"), "6");
}

#[test]
fn lambda_multi_arg() {
    assert_eq!(
        run_python_one("add = lambda a, b: a + b\nprint(add(3, 4))\n"),
        "7"
    );
}

// ══════════════════════════════════════════════════════════════════════════════
// Yield runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn generator_basic_runtime() {
    let out = run_python(
        "def gen():\n    yield 1\n    yield 2\n    yield 3\nfor v in gen():\n    print(v)\n",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

// ══════════════════════════════════════════════════════════════════════════════
// Import varieties
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn import_relative_dot() {
    compile_ok("from . import utils\n");
}

#[test]
fn import_relative_dotdot() {
    compile_ok("from .. import base\n");
}

#[test]
fn import_relative_module() {
    compile_ok("from .utils import helper\n");
}

#[test]
fn import_relative_deep() {
    compile_ok("from ...core import engine\n");
}

#[test]
fn import_star() {
    compile_ok("from math import *\n");
}

#[test]
fn import_parenthesized() {
    compile_ok("from os.path import (\n    join,\n    exists,\n    dirname\n)\n");
}

#[test]
fn import_alias() {
    compile_ok("import numpy as np\n");
}

#[test]
fn import_from_alias() {
    compile_ok("from datetime import datetime as dt\n");
}

#[test]
fn import_multiple() {
    compile_ok("import os, sys, math\n");
}

#[test]
fn import_from_multiple() {
    compile_ok("from os import getcwd, listdir, path\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Edge cases
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn line_continuation_backslash() {
    parse_ok("x = 1 + \\\n    2 + \\\n    3\n");
}

#[test]
fn implicit_continuation_parens() {
    parse_ok("x = (1 +\n     2 +\n     3)\n");
}

#[test]
fn semicolons_multiple_stmts() {
    let out = run_python("x = 1; y = 2; print(x + y)\n");
    assert_eq!(out[0], "3");
}

#[test]
fn single_line_if() {
    assert_eq!(run_python_one("if True: print('yes')\n"), "yes");
}

#[test]
fn single_line_for() {
    compile_ok("for i in range(3): print(i)\n");
}

#[test]
fn single_line_while() {
    compile_ok("while False: pass\n");
}

#[test]
fn single_line_def() {
    assert_eq!(run_python_one("def f(): return 42\nprint(f())\n"), "42");
}

#[test]
fn chained_method_calls() {
    compile_ok("x = 'hello world'.strip().upper().split()\n");
}

#[test]
fn nested_function_calls() {
    compile_ok("x = len(sorted([3, 1, 2]))\n");
}

#[test]
fn pass_in_class() {
    compile_ok("class Empty:\n    pass\n");
}

#[test]
fn pass_in_function() {
    compile_ok("def noop():\n    pass\n");
}

#[test]
fn pass_in_if_else() {
    compile_ok("if True:\n    pass\nelse:\n    pass\n");
}

#[test]
fn del_statement() {
    compile_ok("x = [1, 2, 3]\ndel x[0]\n");
}

#[test]
fn del_variable() {
    compile_ok("x = 42\ndel x\n");
}

#[test]
fn assert_simple() {
    compile_ok("assert True\n");
}

#[test]
fn assert_with_message() {
    compile_ok("assert 1 == 1, 'math is broken'\n");
}
