use crate::helpers::run_python_one;

#[test]
fn decorator_wraps_return_value() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def wrapper(x):\n  return f(x) + 1\n return wrapper\n@deco\ndef inc(n):\n return n\nprint(inc(5))\n"
        ),
        "6"
    );
}

#[test]
fn decorator_preserves_function_name_without_wraps() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def wrapper():\n  return f()\n return wrapper\n@deco\ndef original():\n return 1\nprint(original())\n"
        ),
        "1"
    );
}

#[test]
fn decorator_stacked_outer_inner() {
    assert_eq!(
        run_python_one(
            "def add_a(f):\n def w(x):\n  return 'a' + f(x)\n return w\ndef add_b(f):\n def w(x):\n  return f(x) + 'b'\n return w\n@add_a\n@add_b\ndef mid(x):\n return x\nprint(mid('x'))\n"
        ),
        "axb"
    );
}

#[test]
fn decorator_with_args_factory() {
    assert_eq!(
        run_python_one(
            "def repeat(n):\n def deco(f):\n  def w(x):\n   return f(x) * n\n  return w\n return deco\n@repeat(3)\ndef double(x):\n return x * 2\nprint(double(2))\n"
        ),
        "12"
    );
}

#[test]
fn decorator_class_based_call() {
    assert_eq!(
        run_python_one(
            "class CountCalls:\n def __init__(self, f):\n  self.f = f\n  self.n = 0\n def __call__(self, *a, **k):\n  self.n += 1\n  return self.f(*a, **k)\n@CountCalls\ndef greet():\n return 'hi'\nprint(greet(), greet.n)\n"
        ),
        "hi 2"
    );
}

#[test]
fn decorator_logs_side_effect() {
    assert_eq!(
        run_python_one(
            "log = []\ndef trace(f):\n def w(*a, **k):\n  log.append('call')\n  return f(*a, **k)\n return w\n@trace\ndef add(a, b):\n return a + b\nprint(add(1, 2), log)\n"
        ),
        "3 ['call']"
    );
}

#[test]
fn decorator_on_method() {
    assert_eq!(
        run_python_one(
            "def mark(f):\n def w(self, x):\n  return f(self, x) * 2\n return w\nclass C:\n @mark\n def val(self, x):\n  return x\nprint(C().val(3))\n"
        ),
        "6"
    );
}

#[test]
fn decorator_on_staticmethod() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(x):\n  return f(x) + 1\n return w\nclass U:\n @staticmethod\n @deco\n def inc(x):\n  return x\nprint(U.inc(4))\n"
        ),
        "5"
    );
}

#[test]
fn decorator_on_classmethod() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(cls):\n  return f(cls) + '!' \n return w\nclass A:\n @classmethod\n @deco\n def label(cls):\n  return 'A'\nprint(A.label())\n"
        ),
        "A!"
    );
}

#[test]
fn decorator_passes_through_none() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w():\n  return f()\n return w\n@deco\ndef f():\n return None\nprint(f())\n"
        ),
        "None"
    );
}

#[test]
fn decorator_with_optional_argument_default() {
    assert_eq!(
        run_python_one(
            "def tag(name='t'):\n def deco(f):\n  def w():\n   return name\n  return w\n return deco\n@tag()\ndef f():\n pass\nprint(f())\n"
        ),
        "t"
    );
}

#[test]
fn decorator_factory_custom_tag() {
    assert_eq!(
        run_python_one(
            "def tag(name):\n def deco(f):\n  def w():\n   return name\n  return w\n return deco\n@tag('x')\ndef f():\n pass\nprint(f())\n"
        ),
        "x"
    );
}

#[test]
fn decorator_validates_positive_arg() {
    assert_eq!(
        run_python_one(
            "def positive_only(f):\n def w(x):\n  if x <= 0:\n   raise ValueError('bad')\n  return f(x)\n return w\n@positive_only\ndef sqrt_proxy(x):\n return x\nprint(sqrt_proxy(5))\n"
        ),
        "5"
    );
}

#[test]
fn decorator_catches_exception_returns_default() {
    assert_eq!(
        run_python_one(
            "def safe(f):\n def w():\n  try:\n   return f()\n  except:\n   return 0\n return w\n@safe\ndef boom():\n raise ValueError\nprint(boom())\n"
        ),
        "0"
    );
}

#[test]
fn decorator_timing_sets_flag() {
    assert_eq!(
        run_python_one(
            "seen = []\ndef mark(f):\n def w():\n  seen.append(1)\n  return f()\n return w\n@mark\ndef work():\n return 9\nprint(work(), seen)\n"
        ),
        "9 [1]"
    );
}

#[test]
fn decorator_on_nested_function() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w():\n  return f() + 1\n return w\ndef outer():\n @deco\n def inner():\n  return 1\n return inner\nprint(outer()())\n"
        ),
        "2"
    );
}

#[test]
fn decorator_lambda_wrapped() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n return lambda x: f(x) * 2\ndouble = deco(lambda x: x)\nprint(double(3))\n"
        ),
        "6"
    );
}

#[test]
fn decorator_register_in_dict() {
    assert_eq!(
        run_python_one(
            "REG = {}\ndef register(name):\n def deco(f):\n  REG[name] = f\n  return f\n return deco\n@register('add')\ndef add(a, b):\n return a + b\nprint(REG['add'](2, 3))\n"
        ),
        "5"
    );
}

#[test]
fn decorator_bool_return() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(x):\n  return bool(f(x))\n return w\n@deco\ndef nonzero(x):\n return x\nprint(nonzero(0), nonzero(1))\n"
        ),
        "False True"
    );
}

#[test]
fn decorator_list_return_mutated() {
    assert_eq!(
        run_python_one(
            "def append_marker(f):\n def w():\n  r = f()\n  r.append(9)\n  return r\n return w\n@append_marker\ndef base():\n return [1]\nprint(base())\n"
        ),
        "[1, 9]"
    );
}

#[test]
fn decorator_preserves_args_kwargs() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(*a, **k):\n  return f(*a, **k)\n return w\n@deco\ndef show(a, b=0):\n return a + b\nprint(show(2, b=3))\n"
        ),
        "5"
    );
}

#[test]
fn decorator_on_generator_function() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w():\n  return list(f())\n return w\n@deco\ndef gen():\n yield 1\n yield 2\nprint(gen())\n"
        ),
        "[1, 2]"
    );
}

#[test]
fn decorator_class_decorator_replaces_function() {
    assert_eq!(
        run_python_one(
            "class PlusOne:\n def __init__(self, f):\n  self.f = f\n def __call__(self, x):\n  return self.f(x) + 1\n@PlusOne\ndef inc(x):\n return x\nprint(inc(4))\n"
        ),
        "5"
    );
}

#[test]
fn decorator_property_like_cache() {
    assert_eq!(
        run_python_one(
            "def cache(f):\n stored = {}\n def w(x):\n  if x not in stored:\n   stored[x] = f(x)\n  return stored[x]\n return w\n@cache\ndef sq(x):\n  return x * x\nprint(sq(3), sq(3))\n"
        ),
        "9 9"
    );
}

#[test]
fn decorator_identity_no_op() {
    assert_eq!(
        run_python_one("def identity(f):\n return f\n@identity\ndef f():\n return 7\nprint(f())\n"),
        "7"
    );
}

#[test]
fn decorator_on_function_with_defaults() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(x, y=1):\n  return f(x, y)\n return w\n@deco\ndef add(x, y=1):\n return x + y\nprint(add(5))\n"
        ),
        "6"
    );
}

#[test]
fn decorator_raises_from_wrapped() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w():\n  return f()\n return w\n@deco\ndef bad():\n raise TypeError('t')\ntry:\n bad()\nexcept TypeError:\n print('t')\n"
        ),
        "t"
    );
}

#[test]
fn decorator_on_function_returning_tuple() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w():\n  a, b = f()\n  return a + b\n return w\n@deco\ndef pair():\n return (1, 2)\nprint(pair())\n"
        ),
        "3"
    );
}

#[test]
fn decorator_count_args() {
    assert_eq!(
        run_python_one(
            "def count_args(f):\n def w(*a, **k):\n  return len(a) + len(k)\n return w\n@count_args\ndef dummy(a, b=0):\n pass\nprint(dummy(1, 2, c=3))\n"
        ),
        "3"
    );
}

#[test]
fn decorator_wrap_class_init() {
    assert_eq!(
        run_python_one(
            "def log_init(cls):\n class Wrapped(cls):\n  def __init__(self, *a, **k):\n   super().__init__(*a, **k)\n   self.logged = True\n return Wrapped\n@log_init\nclass C:\n def __init__(self, x):\n  self.x = x\nprint(C(1).logged)\n"
        ),
        "True"
    );
}

#[test]
fn decorator_functools_wraps_name() {
    assert_eq!(
        run_python_one(
            "from functools import wraps\ndef deco(f):\n @wraps(f)\n def w():\n  return f()\n return w\n@deco\ndef named():\n '''doc'''\n return 1\nprint(named.__name__)\n"
        ),
        "named"
    );
}

#[test]
fn decorator_predicate_skip_call() {
    assert_eq!(
        run_python_one(
            "def skip_zero(f):\n def w(x):\n  if x == 0:\n   return None\n  return f(x)\n return w\n@skip_zero\ndef rec(x):\n return x\nprint(skip_zero(rec)(0), skip_zero(rec)(2))\n"
        ),
        "None 2"
    );
}

#[test]
fn decorator_on_async_not_used_sync() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w():\n  return f()\n return w\n@deco\ndef f():\n return 'sync'\nprint(f())\n"
        ),
        "sync"
    );
}

#[test]
fn decorator_triple_stack_order() {
    assert_eq!(
        run_python_one(
            "def d1(f):\n def w(x):\n  return f(x) + 1\n return w\ndef d2(f):\n def w(x):\n  return f(x) * 2\n return w\n@d1\n@d2\ndef f(x):\n return x\nprint(f(3))\n"
        ),
        "7"
    );
}

#[test]
fn decorator_on_function_modifying_string() {
    assert_eq!(
        run_python_one(
            "def upper_result(f):\n def w(s):\n  return f(s).upper()\n return w\n@upper_result\ndef ident(s):\n return s\nprint(ident('hi'))\n"
        ),
        "HI"
    );
}

#[test]
fn decorator_closure_state_increment() {
    assert_eq!(
        run_python_one(
            "def counter(f):\n n = 0\n def w():\n  nonlocal n\n  n += 1\n  return n\n return w\n@counter\ndef tick():\n pass\nprint(tick(), tick())\n"
        ),
        "1 2"
    );
}

#[test]
fn decorator_on_method_preserves_self() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(self, x):\n  return f(self, x) + self.base\n return w\nclass C:\n base = 10\n @deco\n def val(self, x):\n  return x\nprint(C().val(5))\n"
        ),
        "15"
    );
}

#[test]
fn decorator_return_decorator() {
    assert_eq!(
        run_python_one(
            "def make_adder(n):\n def deco(f):\n  def w(x):\n   return f(x) + n\n  return w\n return deco\n@make_adder(5)\ndef id(x):\n return x\nprint(id(1))\n"
        ),
        "6"
    );
}

#[test]
fn decorator_on_function_with_varargs() {
    assert_eq!(
        run_python_one(
            "def sum_args(f):\n def w(*a, **k):\n  return sum(a) + sum(k.values())\n return w\n@sum_args\ndef dummy(*a, **k):\n return 0\nprint(dummy(1, 2, c=3))\n"
        ),
        "6"
    );
}

#[test]
fn decorator_classmethod_alter_return() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(cls):\n  return f(cls) * 2\n return w\nclass A:\n @classmethod\n @deco\n def n(cls):\n  return 3\nprint(A.n())\n"
        ),
        "6"
    );
}

#[test]
fn decorator_staticmethod_chain() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(x):\n  return f(x) + 10\n return w\nclass U:\n @staticmethod\n @deco\n def add1(x):\n  return x + 1\nprint(U.add1(5))\n"
        ),
        "16"
    );
}

#[test]
fn decorator_on_lambda_assigned() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n return lambda x: f(x) + 1\nf = deco(lambda x: x)\nprint(f(8))\n"
        ),
        "9"
    );
}

#[test]
fn decorator_bool_flag_on_function() {
    assert_eq!(
        run_python_one(
            "def flagged(f):\n f.flagged = True\n return f\n@flagged\ndef work():\n return 1\nprint(work.flagged)\n"
        ),
        "True"
    );
}

#[test]
fn decorator_replace_with_constant() {
    assert_eq!(
        run_python_one(
            "def constant(v):\n def deco(f):\n  def w(*a, **k):\n   return v\n  return w\n return deco\n@constant(42)\ndef anything():\n return 0\nprint(anything())\n"
        ),
        "42"
    );
}

#[test]
fn decorator_on_recursive_function() {
    assert_eq!(
        run_python_one(
            "def deco(f):\n def w(n):\n  if n <= 1:\n   return 1\n  return f(n - 1) + f(n - 2)\n return w\n@deco\ndef fib(n):\n return 0\nprint(fib(5))\n"
        ),
        "8"
    );
}
