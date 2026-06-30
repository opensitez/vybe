//! Decorators: functools.wraps, parameterized, class/staticmethod, property, chaining.

use crate::helpers::*;

crate::runtime_case!(
    decorator_plain_wraps,
    "def deco(f):\n def w():\n  return f() + 1\n return w\n@deco\ndef g():\n return 1\nprint(g())\n",
    "2"
);
crate::runtime_case!(
    decorator_with_args_outer,
    "def tag(t):\n def deco(f):\n  def w():\n   return t + f()\n  return w\n return deco\n@tag('x')\ndef g():\n return 'y'\nprint(g())\n",
    "xy"
);
crate::runtime_case!(
    decorator_preserves_name_manual,
    "def deco(f):\n def w():\n  return f()\n w.__name__ = f.__name__\n return w\n@deco\ndef hello():\n pass\nprint(hello.__name__)\n",
    "hello"
);
crate::runtime_case!(
    decorator_class_method,
    "class C:\n @classmethod\n def f(cls):\n  return 'cls'\nprint(C.f())\n",
    "cls"
);
crate::runtime_case!(
    decorator_static_method,
    "class C:\n @staticmethod\n def f():\n  return 9\nprint(C.f())\n",
    "9"
);
crate::runtime_case!(
    decorator_property_getter,
    "class C:\n @property\n def x(self):\n  return 3\nprint(C().x)\n",
    "3"
);
crate::runtime_case!(
    decorator_property_setter,
    "class C:\n def __init__(self):\n  self._v = 0\n @property\n def x(self):\n  return self._v\n @x.setter\n def x(self, v):\n  self._v = v\nc = C()\nc.x = 5\nprint(c.x)\n",
    "5"
);
crate::runtime_case!(
    decorator_stacked_order,
    "def d1(f):\n def w():\n  return 'a' + f()\n return w\ndef d2(f):\n def w():\n  return f() + 'b'\n return w\n@d1\n@d2\ndef g():\n return 'm'\nprint(g())\n",
    "amb"
);
crate::runtime_case!(
    decorator_on_lambda,
    "def deco(f):\n return lambda: f() * 2\ng = deco(lambda: 3)\nprint(g())\n",
    "6"
);
crate::runtime_case!(
    decorator_factory_returns_callable,
    "def repeat(n):\n def deco(f):\n  def w(*a, **k):\n   return f(*a, **k) * n\n  return w\n return deco\n@repeat(3)\ndef add(x, y):\n return x + y\nprint(add(1, 2))\n",
    "9"
);
crate::runtime_case!(
    decorator_mutates_closure,
    "calls = []\ndef log(f):\n def w():\n  calls.append(1)\n  return f()\n return w\n@log\ndef g():\n return 0\ng()\nprint(len(calls))\n",
    "1"
);
crate::runtime_case!(
    decorator_on_method,
    "def deco(f):\n def w(self):\n  return f(self) + 1\n return w\nclass C:\n @deco\n def m(self):\n  return 1\nprint(C().m())\n",
    "2"
);
crate::runtime_case!(
    decorator_builtin_staticmethod,
    "class C:\n @staticmethod\n def add(a, b):\n  return a + b\nprint(C.add(2, 3))\n",
    "5"
);
crate::runtime_case!(
    decorator_builtin_classmethod,
    "class C:\n @classmethod\n def name(cls):\n  return cls.__name__\nprint(C.name())\n",
    "C"
);
crate::runtime_case!(
    decorator_nested_definition,
    "def outer(f):\n def inner():\n  return f()\n return inner\n@outer\ndef base():\n return 7\nprint(base())\n",
    "7"
);
crate::runtime_case!(
    decorator_with_kwargs,
    "def deco(**cfg):\n def wrap(f):\n  def w():\n   return cfg.get('k', 0)\n  return w\n return wrap\n@deco(k=5)\ndef g():\n return 1\nprint(g())\n",
    "5"
);
crate::runtime_case!(
    decorator_class_body,
    "def deco(f):\n return f\n@deco\nclass C:\n x = 1\nprint(C.x)\n",
    "1"
);
crate::runtime_case!(
    decorator_rebind_function,
    "def deco(f):\n return lambda: 99\n@deco\ndef g():\n return 1\nprint(g())\n",
    "99"
);
crate::runtime_case!(
    decorator_access_wrapped_doc,
    "def deco(f):\n def w():\n  '''wrapped'''\n  return f()\n return w\n@deco\ndef g():\n  '''orig'''\n return 0\nprint(callable(g))\n",
    "True"
);
crate::runtime_case!(
    decorator_multiple_methods,
    "def deco(f):\n def w(self):\n  return f(self) * 2\n return w\nclass C:\n @deco\n def a(self):\n  return 1\n @deco\n def b(self):\n  return 2\nprint(C().a(), C().b())\n",
    "2 4"
);
crate::runtime_case!(
    decorator_on_nested_function,
    "def outer():\n @lambda f: f\n def inner():\n  return 3\n return inner\nprint(outer()())\n",
    "3"
);
crate::runtime_case!(
    decorator_generator,
    "def deco(f):\n def w():\n  yield from f()\n return w\n@deco\ndef g():\n yield 1\nprint(list(g()))\n",
    "[1]"
);
crate::runtime_case!(
    decorator_preserves_args,
    "def deco(f):\n def w(x, y):\n  return f(x, y)\n return w\n@deco\ndef add(x, y=0):\n return x + y\nprint(add(2, 3))\n",
    "5"
);
crate::runtime_case!(
    decorator_class_property_deleter,
    "class C:\n def __init__(self):\n  self._x = 1\n @property\n def x(self):\n  return self._x\n @x.deleter\n def x(self):\n  del self._x\nc = C()\ndel c.x\nprint(hasattr(c, '_x'))\n",
    "False"
);
crate::runtime_case!(
    decorator_functools_partial_style,
    "from functools import partial\nadd1 = partial(lambda x, y: x + y, 1)\nprint(add1(2))\n",
    "3"
);
crate::runtime_case!(
    decorator_user_dataclass_like,
    "def dataclass(cls):\n return cls\n@dataclass\nclass P:\n x: int = 0\nprint(P().x)\n",
    "0"
);
crate::runtime_case!(
    decorator_register_pattern,
    "reg = {}\ndef register(name):\n def deco(f):\n  reg[name] = f\n  return f\n return deco\n@register('add')\ndef add(a, b):\n return a + b\nprint(reg['add'](1, 2))\n",
    "3"
);
crate::runtime_case!(
    decorator_timing_pattern,
    "ran = []\ndef track(f):\n def w():\n  ran.append(1)\n  return f()\n return w\n@track\ndef g():\n pass\ng()\nprint(ran)\n",
    "[1]"
);
crate::runtime_case!(
    decorator_bool_flag,
    "enabled = True\ndef optional(f):\n return f if enabled else lambda: None\n@optional\ndef g():\n return 8\nprint(g())\n",
    "8"
);
crate::runtime_case!(
    decorator_simple_identity,
    "def identity(f):\n return f\n@identity\ndef f():\n return 'ok'\nprint(f())\n",
    "ok"
);
crate::runtime_case!(
    decorator_list_comprehension,
    "def deco(f):\n return f\ndef make():\n return [deco(lambda i=i: i) for i in range(3)]\nprint(make()[2]())\n",
    "2"
);
crate::runtime_case!(
    decorator_closure_state,
    "def counter():\n n = 0\n def deco(f):\n  def w():\n   nonlocal n\n   n += 1\n   return n\n  return w\n return deco\n@counter()\ndef g():\n pass\nprint(g())\n",
    "1"
);
crate::runtime_case!(
    decorator_on_init,
    "def log(f):\n return f\nclass C:\n @log\n def __init__(self, v):\n  self.v = v\nprint(C(3).v)\n",
    "3"
);
crate::runtime_case!(
    decorator_wraps_assign_attrs,
    "def deco(f):\n g = lambda: f()\n g.custom = True\n return g\n@deco\ndef h():\n pass\nprint(getattr(h, 'custom', False))\n",
    "True"
);
crate::runtime_case!(
    decorator_method_alter_return,
    "def stringify(f):\n def w(*a, **k):\n  return str(f(*a, **k))\n return w\n@stringify\ndef n():\n return 7\nprint(n())\n",
    "7"
);
crate::runtime_case!(
    decorator_three_deep,
    "def a(f):\n return lambda: f() + 'a'\ndef b(f):\n return lambda: f() + 'b'\ndef c(f):\n return lambda: f() + 'c'\n@c\n@b\n@a\ndef g():\n return ''\nprint(g())\n",
    "cba"
);
crate::runtime_case!(
    decorator_on_lambda_assigned,
    "deco = lambda f: (lambda: f() + 1)\nf = deco(lambda: 1)\nprint(f())\n",
    "2"
);
crate::runtime_case!(
    decorator_exception_swallow,
    "def swallow(f):\n def w():\n  try:\n   return f()\n  except ValueError:\n   return 0\n return w\n@swallow\ndef g():\n raise ValueError()\nprint(g())\n",
    "0"
);
crate::runtime_case!(
    decorator_with_self_parameter,
    "def deco(f):\n def w(self, x):\n  return f(self, x) * 2\n return w\nclass C:\n @deco\n def m(self, x):\n  return x\nprint(C().m(4))\n",
    "8"
);
crate::runtime_case!(
    decorator_import_functools_wraps,
    "from functools import wraps\ndef deco(f):\n @wraps(f)\n def w():\n  return f()\n return w\n@deco\ndef g():\n  '''doc'''\n pass\nprint(g.__name__)\n",
    "g"
);
crate::runtime_case!(
    decorator_lru_cache_pattern,
    "from functools import lru_cache\n@lru_cache(maxsize=None)\ndef fib(n):\n return n if n < 2 else fib(n-1) + fib(n-2)\nprint(fib(5))\n",
    "5"
);
crate::runtime_case!(
    decorator_total_ordering,
    "from functools import total_ordering\n@total_ordering\nclass C:\n def __init__(self, v):\n  self.v = v\n def __eq__(self, o):\n  return self.v == o.v\n def __lt__(self, o):\n  return self.v < o.v\nprint(C(1) < C(2))\n",
    "True"
);
crate::runtime_case!(
    decorator_singledispatch,
    "from functools import singledispatch\n@singledispatch\ndef f(x):\n return 'd'\n@f.register(int)\ndef _i(x):\n return 'i'\nprint(f(1))\n",
    "i"
);

crate::compile_case!(decorator_async, "def deco(f):\n return f\n@deco\nasync def ag():\n return 1\n");
crate::compile_case!(decorator_classmethod_property, "class C:\n @classmethod\n @property\n def x(cls):\n  return 1\n");
crate::compile_case!(decorator_abstractmethod, "from abc import abstractmethod\nclass B:\n @abstractmethod\n def m(self):\n  pass\n");
crate::compile_case!(decorator_cached_property, "class C:\n @property\n def x(self):\n  return 1\n");
crate::compile_case!(decorator_parametrized_stack, "def a(x):\n def deco(f):\n  return f\n return deco\n@a(1)\n@a(2)\ndef f():\n pass\n");
