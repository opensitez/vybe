//! Closure capture, late binding, nonlocal, cell variables.

crate::runtime_case!(
    closure_capture_read,
    "def outer():\n x = 1\n def inner():\n  return x\n return inner\nprint(outer()())\n",
    "1"
);
crate::runtime_case!(
    closure_capture_write_nonlocal,
    "def outer():\n x = 1\n def inner():\n  nonlocal x\n  x = 2\n  return x\n inner()\n return x\nprint(outer())\n",
    "2"
);
crate::runtime_case!(
    closure_late_binding,
    "funcs = []\nfor i in range(3):\n funcs.append(lambda: i)\nprint(funcs[2]())\n",
    "2"
);
crate::runtime_case!(
    closure_default_arg_fix,
    "funcs = []\nfor i in range(3):\n funcs.append(lambda i=i: i)\nprint(funcs[1]())\n",
    "1"
);
crate::runtime_case!(
    closure_nested_levels,
    "def a():\n x = 1\n def b():\n  y = 2\n  def c():\n   return x + y\n  return c\n return b()()\nprint(a())\n",
    "3"
);
crate::runtime_case!(
    closure_factory,
    "def make_adder(n):\n return lambda x: x + n\nprint(make_adder(5)(3))\n",
    "8"
);
crate::runtime_case!(
    closure_in_loop,
    "def make():\n out = []\n for i in range(3):\n  out.append(lambda i=i: i * 2)\n return out\nprint(make()[2]())\n",
    "4"
);
crate::runtime_case!(
    closure_global_shadow,
    "x = 10\ndef outer():\n def inner():\n  return x\n return inner\nprint(outer()())\n",
    "10"
);
crate::runtime_case!(
    closure_global_write,
    "x = 1\ndef outer():\n global x\n def inner():\n  global x\n  x = 2\n inner()\nouter()\nprint(x)\n",
    "2"
);
crate::runtime_case!(
    closure_builtin,
    "def outer():\n return lambda: len([1, 2])\nprint(outer()())\n",
    "2"
);
crate::runtime_case!(
    closure_as_default,
    "def f(g=lambda: 1):\n return g()\nprint(f())\n",
    "1"
);
crate::runtime_case!(
    closure_return_closure,
    "def counter():\n n = 0\n def inc():\n  nonlocal n\n  n += 1\n  return n\n return inc\nprint(counter()())\n",
    "1"
);
crate::runtime_case!(
    closure_shared_state,
    "def counter():\n n = 0\n def inc():\n  nonlocal n\n  n += 1\n  return n\n return inc\nc = counter()\nprint(c(), c())\n",
    "1 2"
);
crate::runtime_case!(
    closure_cell_independent,
    "def make():\n return lambda: 1\na = make()\nb = make()\nprint(a(), b())\n",
    "1\n1"
);
crate::runtime_case!(
    closure_truthy,
    "def outer():\n return lambda: True\nprint(bool(outer()))\n",
    "True"
);
crate::runtime_case!(
    closure_callable,
    "def outer():\n return lambda: None\nprint(callable(outer()))\n",
    "True"
);
crate::runtime_case!(
    closure_name,
    "def outer():\n return lambda: 1\nprint(outer().__name__)\n",
    "<lambda>"
);
crate::runtime_case!(
    nested_def_name,
    "def outer():\n def inner():\n  pass\n return inner.__name__\nprint(outer())\n",
    "inner"
);
crate::runtime_case!(
    closure_freevars,
    "def outer(x):\n def inner():\n  return x\n return inner\nprint(outer(7).__code__.co_freevars)\n",
    "('x',)"
);
crate::runtime_case!(
    closure_nonlocal_shadow,
    "def outer():\n x = 1\n def middle():\n  x = 2\n  def inner():\n   nonlocal x\n   return x\n  return inner()\n return middle()\nprint(outer())\n",
    "2"
);
crate::runtime_case!(
    closure_decorator,
    "def deco(f):\n return lambda: f() + 1\ndef g():\n return 1\nprint(deco(g)())\n",
    "2"
);
crate::runtime_case!(
    closure_in_class_method,
    "class C:\n def m(self):\n  x = 1\n  return lambda: x\nprint(C().m()())\n",
    "1"
);
crate::runtime_case!(
    closure_list_comp,
    "n = 2\nprint([lambda: n for _ in range(1)][0]())\n",
    "2"
);
crate::runtime_case!(
    closure_generator,
    "def outer():\n x = 1\n def gen():\n  yield x\n return gen\nprint(list(outer()()))\n",
    "[1]"
);
crate::runtime_case!(
    closure_exception_handler,
    "def outer():\n msg = 'err'\n def inner():\n  try:\n   raise ValueError(msg)\n  except ValueError as e:\n   return str(e)\n return inner\nprint(outer()())\n",
    "err"
);
crate::runtime_case!(
    closure_del_closure,
    "def outer():\n x = 1\n def inner():\n  return x\n return inner\nf = outer()\nprint(f())\n",
    "1"
);
crate::runtime_case!(
    closure_recursive,
    "def make_fact():\n def fact(n):\n  return 1 if n <= 1 else n * fact(n - 1)\n return fact\nprint(make_fact()(5))\n",
    "120"
);
crate::runtime_case!(
    closure_partial_apply,
    "def add(a, b):\n return a + b\ndef bind_a(a):\n return lambda b: add(a, b)\nprint(bind_a(3)(4))\n",
    "7"
);
crate::runtime_case!(
    closure_compare,
    "def outer():\n return lambda x: x > 0\nprint(outer()(5))\n",
    "True"
);
crate::runtime_case!(
    closure_bool_context,
    "def outer():\n return lambda: 0\nprint(bool(outer()()))\n",
    "False"
);
crate::runtime_case!(
    closure_in_dict,
    "def outer():\n return {'f': lambda: 9}\nprint(outer()['f']())\n",
    "9"
);
crate::runtime_case!(
    closure_in_tuple,
    "def outer():\n return (lambda: 1,)\nprint(outer()[0]())\n",
    "1"
);
crate::runtime_case!(
    closure_nonlocal_two_vars,
    "def outer():\n a = 1\n b = 2\n def inner():\n  nonlocal a, b\n  a += 1\n  b += 1\n  return a + b\n return inner\nprint(outer()())\n",
    "5"
);
crate::runtime_case!(
    closure_read_before_assign,
    "def outer():\n x = 1\n def inner():\n  return x\n return inner\nprint(outer()())\n",
    "1"
);
crate::runtime_case!(
    closure_enclosing_name,
    "def outer():\n def inner():\n  return inner.__name__\n return inner\nprint(outer()())\n",
    "inner"
);
crate::runtime_case!(
    closure_walrus,
    "def outer():\n def inner():\n  if (x := 5) > 0:\n   return x\n return inner\nprint(outer()())\n",
    "5"
);
crate::runtime_case!(
    closure_match,
    "def outer():\n def inner(x):\n  match x:\n   case 1:\n    return 'one'\n   case _:\n    return 'other'\n return inner\nprint(outer()(1))\n",
    "one"
);
crate::runtime_case!(
    closure_async_syntax,
    "def outer():\n async def inner():\n  return 1\n return inner\nprint(callable(outer()))\n",
    "True"
);
crate::runtime_case!(
    closure_class_cell,
    "def outer():\n class C:\n  x = 1\n return C\nprint(outer().x)\n",
    "1"
);
crate::runtime_case!(
    closure_default_mutable,
    "def outer():\n cache = []\n def inner(v):\n  cache.append(v)\n  return len(cache)\n return inner\nf = outer()\nprint(f(1), f(2))\n",
    "1 2"
);
crate::runtime_case!(
    closure_id_stable,
    "def outer():\n return lambda: id(1)\na = outer()\nb = outer()\nprint(a() == b())\n",
    "True"
);
crate::runtime_case!(
    closure_isinstance,
    "def outer():\n return lambda: isinstance(1, int)\nprint(outer()())\n",
    "True"
);
crate::runtime_case!(
    closure_type_check,
    "def outer():\n return lambda: type(1).__name__\nprint(outer()())\n",
    "int"
);
crate::runtime_case!(
    closure_higher_order,
    "def apply(f):\n return f(10)\ndef make():\n return lambda x: x + 1\nprint(apply(make()))\n",
    "11"
);
crate::runtime_case!(
    closure_filter_map,
    "def outer():\n n = 5\n return lambda x: x < n\nprint(list(filter(outer(), [1, 6, 3])))\n",
    "[1, 3]"
);

crate::compile_case!(
    closure_nonlocal_error,
    "def outer():\n def inner():\n  nonlocal x\n"
);
crate::compile_case!(
    closure_global_nonlocal_mix,
    "x = 1\ndef outer():\n def inner():\n  global x\n  x = 2\n"
);
crate::compile_case!(
    closure_class_scope,
    "def outer():\n class C:\n  def m(self):\n   return x\n"
);
crate::compile_case!(
    closure_annotations,
    "def outer():\n x: int = 1\n def inner() -> int:\n  return x\n"
);
crate::compile_case!(
    closure_yield_from,
    "def outer():\n def inner():\n  yield from range(2)\n return inner\n"
);
