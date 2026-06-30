//! Function signatures: /, *, **, defaults, kw-only — runtime calls.

crate::runtime_case!(
    positional_only_call,
    "def f(a, b, /):\n return a + b\nprint(f(1, 2))\n",
    "3"
);
crate::runtime_case!(
    positional_only_three,
    "def f(a, b, c, /):\n return a + b + c\nprint(f(1, 2, 3))\n",
    "6"
);
crate::runtime_case!(
    positional_only_with_kw,
    "def f(a, b, /, c):\n return a + b + c\nprint(f(1, 2, c=3))\n",
    "6"
);
crate::runtime_case!(
    positional_only_kwonly,
    "def f(a, /, *, b):\n return a + b\nprint(f(1, b=2))\n",
    "3"
);
crate::runtime_case!(
    bare_star_kwonly,
    "def f(a, *, b, c):\n return a + b + c\nprint(f(1, b=2, c=3))\n",
    "6"
);
crate::runtime_case!(
    kwonly_defaults,
    "def f(*, a=1, b=2):\n return a + b\nprint(f())\n",
    "3"
);
crate::runtime_case!(
    positional_defaults,
    "def f(a, b=10):\n return a + b\nprint(f(5))\n",
    "15"
);
crate::runtime_case!(
    varargs_sum,
    "def f(*args):\n return sum(args)\nprint(f(1, 2, 3))\n",
    "6"
);
crate::runtime_case!(
    kwargs_len,
    "def f(**kwargs):\n return len(kwargs)\nprint(f(a=1, b=2))\n",
    "2"
);
crate::runtime_case!(
    mixed_all,
    "def f(a, b=2, /, c=3, *args, d, **kwargs):\n return (a, b, c, args, d, kwargs)\nprint(f(1, 2, 4, d=5, e=6)[0])\n",
    "1"
);
crate::runtime_case!(
    lambda_positional,
    "f = lambda x, y: x * y\nprint(f(3, 4))\n",
    "12"
);
crate::runtime_case!(
    lambda_defaults,
    "f = lambda x, y=2: x + y\nprint(f(3))\n",
    "5"
);
crate::runtime_case!(
    lambda_varargs,
    "f = lambda *a: len(a)\nprint(f(1, 2, 3))\n",
    "3"
);
crate::runtime_case!(
    lambda_kwargs,
    "f = lambda **k: sorted(k)\nprint(f(b=2, a=1))\n",
    "['a', 'b']"
);
crate::runtime_case!(
    def_annotations,
    "def f(x: int) -> int:\n return x + 1\nprint(f(1))\n",
    "2"
);
crate::runtime_case!(
    def_annotation_str,
    "def f(x: 'int') -> 'int':\n return x\nprint(f(5))\n",
    "5"
);
crate::runtime_case!(
    keyword_only_error,
    "def f(a, *, b):\n return a + b\ntry:\n f(1, 2)\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    too_many_positional,
    "def f(a, b):\n return a + b\ntry:\n f(1, 2, 3)\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    missing_required,
    "def f(a, b):\n return a + b\ntry:\n f(1)\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    unexpected_keyword,
    "def f(a):\n return a\ntry:\n f(a=1, b=2)\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    unpack_positional,
    "def f(a, b, c):\n return a + b + c\nprint(f(*[1, 2, 3]))\n",
    "6"
);
crate::runtime_case!(
    unpack_keyword,
    "def f(a, b):\n return a + b\nprint(f(**{'a': 1, 'b': 2}))\n",
    "3"
);
crate::runtime_case!(
    unpack_mixed,
    "def f(a, b, c):\n return a + b + c\nprint(f(1, *[2], c=3))\n",
    "6"
);
crate::runtime_case!(
    default_mutable_none,
    "def f(x=None):\n if x is None:\n  x = []\n x.append(1)\n return len(x)\nprint(f())\n",
    "1"
);
crate::runtime_case!(
    nested_defaults,
    "def f(a, b=factory if False else 10):\n return a + b\ndef g(a, b=10):\n return a + b\nprint(g(5))\n",
    "15"
);
crate::runtime_case!(
    method_self_implicit,
    "class C:\n def m(self, x):\n  return x + 1\nprint(C().m(4))\n",
    "5"
);
crate::runtime_case!(
    classmethod_no_self,
    "class C:\n @classmethod\n def m(cls, x):\n  return x\nprint(C.m(9))\n",
    "9"
);
crate::runtime_case!(
    staticmethod_no_self,
    "class C:\n @staticmethod\n def m(x):\n  return x * 2\nprint(C.m(4))\n",
    "8"
);
crate::runtime_case!(
    partial_application,
    "from functools import partial\ndef add(a, b, c):\n return a + b + c\nprint(partial(add, 1)(2, 3))\n",
    "6"
);
crate::runtime_case!(
    inspect_signature_params,
    "import inspect\ndef f(a, b=1, *, c):\n pass\nprint(len(inspect.signature(f).parameters))\n",
    "3"
);
crate::runtime_case!(
    __defaults__tuple,
    "def f(a, b=1):\n pass\nprint(f.__defaults__)\n",
    "(1,)"
);
crate::runtime_case!(
    __kwdefaults__dict,
    "def f(*, a=1):\n pass\nprint(f.__kwdefaults__)\n",
    "{'a': 1}"
);
crate::runtime_case!(
    __annotations__dict,
    "def f(x: int) -> str:\n pass\nprint(f.__annotations__['x'].__name__)\n",
    "int"
);
crate::runtime_case!(
    positional_only_name_error,
    "def f(a, /, b):\n return a + b\ntry:\n f(a=1, b=2)\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    keyword_after_positional,
    "def f(a, b, c):\n return a + b + c\nprint(f(1, c=3, b=2))\n",
    "6"
);
crate::runtime_case!(
    empty_varargs,
    "def f(*args):\n return args\nprint(f())\n",
    "()"
);
crate::runtime_case!(
    empty_kwargs,
    "def f(**kwargs):\n return kwargs\nprint(f())\n",
    "{}"
);
crate::runtime_case!(
    nested_function_args,
    "def outer(x):\n def inner(y):\n  return x + y\n return inner\nprint(outer(1)(2))\n",
    "3"
);
crate::runtime_case!(
    generator_function_yield,
    "def g():\n yield 1\nprint(list(g()))\n",
    "[1]"
);
crate::runtime_case!(
    recursive_factorial,
    "def fact(n):\n return 1 if n <= 1 else n * fact(n - 1)\nprint(fact(5))\n",
    "120"
);
crate::runtime_case!(
    call_with_none_args,
    "def f(a=None):\n return a\nprint(f(None))\n",
    "None"
);
crate::runtime_case!(
    call_with_bool_kwargs,
    "def f(**k):\n return k.get('x', False)\nprint(f(x=True))\n",
    "True"
);
crate::runtime_case!(
    positional_only_slash_only,
    "def f(a, /):\n return a\nprint(f(7))\n",
    "7"
);
crate::runtime_case!(
    dual_slash_error,
    "def f(a, b, /, c):\n return a + b + c\nprint(f(1, 2, 3))\n",
    "6"
);
crate::runtime_case!(
    keyword_varargs_conflict,
    "def f(*args, kw):\n return kw\nprint(f(1, 2, kw=3))\n",
    "3"
);
crate::runtime_case!(
    function_name_qualname,
    "def f():\n pass\nprint(f.__name__)\n",
    "f"
);
crate::runtime_case!(
    nested_qualname,
    "def outer():\n def inner():\n  return inner.__name__\n return inner\nprint(outer()())\n",
    "inner"
);

crate::compile_case!(positional_only_after_slash_error, "def f(a, /, /, b): pass\n");
crate::compile_case!(async_def_signature, "async def f(a, /, *, b): pass\n");
crate::compile_case!(type_params_pep695, "def f[T](x: T) -> T: return x\n");
crate::compile_case!(keyword_only_before_star, "def f(*, a, b): pass\n");
crate::compile_case!(inspect_bound_arguments, "import inspect\ndef f(a, b=1): pass\ninspect.signature(f).bind(1)\n");
