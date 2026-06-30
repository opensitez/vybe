//! eval, exec, compile builtins and code objects.

crate::runtime_case!(
    eval_literal_int,
    "print(eval('1 + 2'))\n",
    "3"
);
crate::runtime_case!(
    eval_literal_list,
    "print(eval('[1, 2, 3]'))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    eval_with_globals,
    "print(eval('x + 1', {'x': 5}))\n",
    "6"
);
crate::runtime_case!(
    eval_with_locals,
    "print(eval('x', {}, {'x': 9}))\n",
    "9"
);
crate::runtime_case!(
    exec_assign,
    "ns = {}\nexec('y = 7', ns)\nprint(ns['y'])\n",
    "7"
);
crate::runtime_case!(
    exec_multiline,
    "ns = {}\nexec('a = 1\\nb = 2', ns)\nprint(ns['b'])\n",
    "2"
);
crate::runtime_case!(
    compile_eval_mode,
    "c = compile('2 + 3', '<string>', 'eval')\nprint(eval(c))\n",
    "5"
);
crate::runtime_case!(
    compile_exec_mode,
    "c = compile('z = 4', '<string>', 'exec')\nns = {}\nexec(c, ns)\nprint(ns['z'])\n",
    "4"
);
crate::runtime_case!(
    compile_single_mode,
    "c = compile('print(9)', '<string>', 'single')\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    eval_expr_name,
    "x = 10\nprint(eval('x'))\n",
    "10"
);
crate::runtime_case!(
    eval_builtin_call,
    "print(eval('len([1,2,3])'))\n",
    "3"
);
crate::runtime_case!(
    exec_function_def,
    "ns = {}\nexec('def f(): return 42', ns)\nprint(ns['f']())\n",
    "42"
);
crate::runtime_case!(
    exec_class_def,
    "ns = {}\nexec('class C:\\n x = 1', ns)\nprint(ns['C'].x)\n",
    "1"
);
crate::runtime_case!(
    compile_co_name,
    "c = compile('1', '<s>', 'eval')\nprint(c.co_name)\n",
    "<module>"
);
crate::runtime_case!(
    compile_co_consts,
    "c = compile('1 + 2', '<s>', 'eval')\nprint(1 in c.co_consts)\n",
    "True"
);
crate::runtime_case!(
    eval_dict_literal,
    "print(eval('{\"a\": 1}'))\n",
    "{'a': 1}"
);
crate::runtime_case!(
    eval_bool_ops,
    "print(eval('True and False'))\n",
    "False"
);
crate::runtime_case!(
    eval_comparison,
    "print(eval('3 < 5'))\n",
    "True"
);
crate::runtime_case!(
    exec_import,
    "ns = {}\nexec('import math', ns)\nprint(ns['math'].sqrt(9))\n",
    "3.0"
);
crate::runtime_case!(
    eval_lambda,
    "print(eval('(lambda x: x + 1)(5)'))\n",
    "6"
);
crate::runtime_case!(
    compile_ast_mode,
    "c = compile('1+1', '<s>', 'eval', flags=0)\nprint(eval(c))\n",
    "2"
);
crate::runtime_case!(
    eval_nested,
    "print(eval('eval(\"1+1\")'))\n",
    "2"
);
crate::runtime_case!(
    exec_for_loop,
    "ns = {'out': []}\nexec('for i in range(3): out.append(i)', ns)\nprint(ns['out'])\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    exec_if_else,
    "ns = {}\nexec('x = 1 if True else 0', ns)\nprint(ns['x'])\n",
    "1"
);
crate::runtime_case!(
    eval_tuple,
    "print(eval('(1, 2)'))\n",
    "(1, 2)"
);
crate::runtime_case!(
    eval_set,
    "print(sorted(eval('{1, 2, 3}')))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    compile_filename,
    "c = compile('1', 'file.py', 'eval')\nprint(c.co_filename)\n",
    "file.py"
);
crate::runtime_case!(
    eval_string_concat,
    "print(eval(\"'a' + 'b'\"))\n",
    "ab"
);
crate::runtime_case!(
    exec_raise_caught,
    "ns = {}\ntry:\n exec('raise ValueError(\"e\")', ns)\n print('ok')\nexcept ValueError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    eval_none,
    "print(eval('None'))\n",
    "None"
);
crate::runtime_case!(
    eval_float,
    "print(eval('3.14'))\n",
    "3.14"
);
crate::runtime_case!(
    compile_dont_inherit,
    "c = compile('1', '<s>', 'eval', dont_inherit=True)\nprint(eval(c))\n",
    "1"
);
crate::runtime_case!(
    exec_delete,
    "ns = {'x': 1}\nexec('del x', ns)\nprint('x' in ns)\n",
    "False"
);
crate::runtime_case!(
    eval_list_comp,
    "print(eval('[x for x in range(3)]'))\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    eval_dict_comp,
    "print(eval('{x: x for x in range(2)}'))\n",
    "{0: 0, 1: 1}"
);
crate::runtime_case!(
    compile_flags,
    "import ast\nc = compile('1+1', '<s>', 'eval', flags=ast.PyCF_ONLY_AST)\nprint(type(c).__name__)\n",
    "Module"
);
crate::runtime_case!(
    eval_builtin_enumerate,
    "print(eval('list(enumerate([\"a\"]))'))\n",
    "[(0, 'a')]"
);
crate::runtime_case!(
    exec_while_loop,
    "ns = {'n': 0, 's': 0}\nexec('while n < 3:\\n s += n; n += 1', ns)\nprint(ns['s'])\n",
    "3"
);
crate::runtime_case!(
    eval_getattr,
    "print(eval('(1).__class__.__name__'))\n",
    "int"
);
crate::runtime_case!(
    compile_optimize,
    "c = compile('1', '<s>', 'eval', optimize=2)\nprint(eval(c))\n",
    "1"
);
crate::runtime_case!(
    exec_with_statement,
    "ns = {}\nexec('class CM:\\n def __enter__(self): return 1\\n def __exit__(self, *a): pass\\nwith CM() as v: x = v', ns)\nprint(ns['x'])\n",
    "1"
);
crate::runtime_case!(
    eval_ternary,
    "print(eval('1 if 2 > 1 else 0'))\n",
    "1"
);
crate::runtime_case!(
    exec_try_except,
    "ns = {}\nexec('try:\\n 1/0\\nexcept ZeroDivisionError:\\n err = 1', ns)\nprint(ns['err'])\n",
    "1"
);
crate::runtime_case!(
    eval_bytes_literal,
    "print(eval(\"b'hi'\"))\n",
    "b'hi'"
);
crate::runtime_case!(
    compile_malformed_raises,
    "try:\n compile('1 +', '<s>', 'eval')\n print('ok')\nexcept SyntaxError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    eval_name_error,
    "try:\n eval('undefined_xyz')\n print('ok')\nexcept NameError:\n print('err')\n",
    "err"
);

crate::compile_case!(eval_ast_mode, "import ast\ncode = compile('1+1', '<s>', 'eval', flags=ast.PyCF_ONLY_AST)\n");
crate::compile_case!(exec_globals_locals, "exec('pass', {}, {})\n");
crate::compile_case!(compile_annotate, "compile('x: int = 1', '<s>', 'exec')\n");
crate::compile_case!(eval_restricted, "eval('1', {'__builtins__': {}}, {})\n");
crate::compile_case!(types_code_type, "import types\ntypes.CodeType\n");
