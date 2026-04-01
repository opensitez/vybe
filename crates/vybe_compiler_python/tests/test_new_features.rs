//! Tests for the 10 new Python features.
//! Each test compiles the source to verify the compiler handles the syntax correctly.

use vybe_parser_python::parse;
use vybe_compiler_python::Compiler;

fn compile_ok(src: &str) {
    let module = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&module);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── 1. Slice with step ──────────────────────────────────────

#[test] fn slice_step_reverse() { compile_ok("x = [1,2,3,4,5]\ny = x[::-1]\n"); }
#[test] fn slice_step_every_other() { compile_ok("x = [1,2,3,4,5,6]\ny = x[::2]\n"); }
#[test] fn slice_step_with_bounds() { compile_ok("x = [1,2,3,4,5]\ny = x[1:4:2]\n"); }
#[test] fn slice_step_negative() { compile_ok("x = [1,2,3,4,5]\ny = x[4:1:-1]\n"); }
#[test] fn slice_no_step() { compile_ok("x = [1,2,3,4,5]\ny = x[1:3]\n"); }
#[test] fn slice_step_string() { compile_ok("s = 'hello'\nr = s[::-1]\n"); }
#[test] fn slice_step_only() { compile_ok("x = [1,2,3,4,5]\ny = x[::3]\n"); }

// ── 2. String repetition ────────────────────────────────────

#[test] fn str_repeat_basic() { compile_ok("x = '=' * 40\n"); }
#[test] fn str_repeat_int_first() { compile_ok("x = 3 * 'abc'\n"); }
#[test] fn str_repeat_variable() { compile_ok("n = 5\nx = '-' * n\n"); }
#[test] fn str_repeat_in_expr() { compile_ok("print('*' * 10 + ' hello ' + '*' * 10)\n"); }
#[test] fn int_multiply_still_works() { compile_ok("x = 6 * 7\n"); }
#[test] fn float_multiply_still_works() { compile_ok("x = 3.14 * 2.0\n"); }
#[test] fn str_repeat_augmented() { compile_ok("s = 'x'\ns *= 5\n"); }

// ── 3. del statement ────────────────────────────────────────

#[test] fn del_variable() { compile_ok("x = 10\ndel x\n"); }
#[test] fn del_dict_key() { compile_ok("d = {'a': 1, 'b': 2}\ndel d['a']\n"); }
#[test] fn del_list_index() { compile_ok("lst = [1, 2, 3]\ndel lst[0]\n"); }
#[test] fn del_attribute() { compile_ok("class C:\n    pass\nc = C()\nc.x = 10\ndel c.x\n"); }
#[test] fn del_multiple() { compile_ok("a = 1\nb = 2\ndel a, b\n"); }

// ── 4. Match/case ───────────────────────────────────────────

#[test] fn match_basic_value() {
    compile_ok("x = 42\nmatch x:\n    case 1:\n        print('one')\n    case 42:\n        print('forty-two')\n    case _:\n        print('other')\n");
}
#[test] fn match_wildcard() {
    compile_ok("match 'hello':\n    case _:\n        print('anything')\n");
}
#[test] fn match_or_pattern() {
    compile_ok("x = 2\nmatch x:\n    case 1 | 2 | 3:\n        print('small')\n    case _:\n        print('big')\n");
}
#[test]
fn match_with_guard() {
    compile_ok("x = 10\nmatch x:\n    case n if n > 5:\n        print('big')\n    case _:\n        print('small')\n");
}
#[test] fn match_none() {
    compile_ok("x = None\nmatch x:\n    case None:\n        print('none')\n    case _:\n        print('other')\n");
}
#[test] fn match_string() {
    compile_ok("cmd = 'quit'\nmatch cmd:\n    case 'quit' | 'exit':\n        print('bye')\n    case 'help':\n        print('help')\n    case _:\n        pass\n");
}

// ── 5. Generators (yield) ───────────────────────────────────

#[test] fn yield_basic() { compile_ok("def gen():\n    yield 1\n    yield 2\n    yield 3\n"); }
#[test] fn yield_no_value() { compile_ok("def gen():\n    yield\n"); }
#[test] fn yield_in_loop() { compile_ok("def count_up(n):\n    i = 0\n    while i < n:\n        yield i\n        i += 1\n"); }
#[test] fn yield_from() { compile_ok("def chain(a, b):\n    yield from a\n    yield from b\n"); }

// ── 6. For/while else ───────────────────────────────────────

#[test] fn for_else_no_break() {
    compile_ok("for x in [1, 2, 3]:\n    if x == 5:\n        break\nelse:\n    print('no five')\n");
}
#[test] fn for_else_with_break() {
    compile_ok("for x in [1, 2, 3]:\n    if x == 2:\n        break\nelse:\n    print('unreachable')\n");
}
#[test] fn while_else() {
    compile_ok("i = 0\nwhile i < 5:\n    i += 1\nelse:\n    print('done')\n");
}
#[test] fn while_else_with_break() {
    compile_ok("i = 0\nwhile i < 10:\n    if i == 3:\n        break\n    i += 1\nelse:\n    print('completed')\n");
}

// ── 7. Multiple inheritance ─────────────────────────────────

#[test] fn single_inheritance() {
    compile_ok("class Animal:\n    def speak(self):\n        return 'generic'\n\nclass Dog(Animal):\n    def speak(self):\n        return 'woof'\n");
}
#[test] fn multiple_inheritance() {
    compile_ok("class A:\n    def method_a(self):\n        return 'a'\n\nclass B:\n    def method_b(self):\n        return 'b'\n\nclass C(A, B):\n    pass\n");
}
#[test] fn diamond_inheritance() {
    compile_ok("class Base:\n    pass\nclass Left(Base):\n    pass\nclass Right(Base):\n    pass\nclass Child(Left, Right):\n    pass\n");
}

// ── 8. Extended unpacking in literals ────────────────────────

#[test] fn list_unpack_star() { compile_ok("a = [2, 3]\nb = [1, *a, 4]\n"); }
#[test] fn list_unpack_multiple() { compile_ok("x = [1, 2]\ny = [3, 4]\nz = [*x, *y]\n"); }
#[test] fn tuple_unpack_star() { compile_ok("a = (1, 2)\nb = (*a, 3, 4)\n"); }
#[test] fn list_unpack_empty() { compile_ok("a = []\nb = [1, *a, 2]\n"); }
#[test] fn list_no_star_still_works() { compile_ok("a = [1, 2, 3]\n"); }

// ── 9. Slice assignment ─────────────────────────────────────

#[test] fn slice_assign_basic() { compile_ok("a = [1, 2, 3, 4, 5]\na[1:3] = [10, 20]\n"); }
#[test] fn slice_assign_empty() { compile_ok("a = [1, 2, 3]\na[1:1] = [10, 20]\n"); }
#[test] fn slice_assign_delete() { compile_ok("a = [1, 2, 3, 4, 5]\na[1:3] = []\n"); }
#[test] fn slice_assign_full() { compile_ok("a = [1, 2, 3]\na[:] = [4, 5, 6]\n"); }

// ── 10. del for dict keys and list indices (covered by #3) ──

#[test] fn del_nested_dict() { compile_ok("d = {'a': {'b': 1}}\ndel d['a']\n"); }
#[test] fn del_dict_variable_key() { compile_ok("d = {'x': 1}\nk = 'x'\ndel d[k]\n"); }

// ── Method dispatch: user methods override builtins ──────────

#[test] fn user_method_named_get() {
    compile_ok("class C:\n    def get(self):\n        return 42\nc = C()\nprint(c.get())\n");
}
#[test] fn user_method_named_keys() {
    compile_ok("class C:\n    def keys(self):\n        return ['a']\nc = C()\nprint(c.keys())\n");
}
#[test] fn user_method_named_pop() {
    compile_ok("class C:\n    def pop(self):\n        return 'popped'\nc = C()\nprint(c.pop())\n");
}
#[test] fn user_method_named_append() {
    compile_ok("class C:\n    def append(self, x):\n        print(x)\nc = C()\nc.append(42)\n");
}
#[test] fn builtin_list_append_still_works() {
    compile_ok("x = [1, 2]\nx.append(3)\nprint(x)\n");
}
#[test] fn builtin_dict_get_still_works() {
    compile_ok("d = {'a': 1}\nprint(d.get('a'))\n");
}
#[test] fn builtin_dict_keys_still_works() {
    compile_ok("d = {'a': 1, 'b': 2}\nprint(d.keys())\n");
}

// ── map / filter / iter / next ─────────────────────────────

#[test] fn map_basic() {
    compile_ok("result = list(map(lambda x: x * 2, [1, 2, 3]))\nprint(result)\n");
}
#[test] fn map_with_named_func() {
    compile_ok("def double(x):\n    return x * 2\nresult = map(double, [1, 2, 3])\n");
}
#[test] fn filter_basic() {
    compile_ok("result = list(filter(lambda x: x > 0, [-1, 0, 1, 2]))\nprint(result)\n");
}
#[test] fn filter_with_named_func() {
    compile_ok("def is_even(x):\n    return x % 2 == 0\nresult = filter(is_even, [1, 2, 3, 4])\n");
}
#[test] fn iter_passthrough() {
    compile_ok("it = iter([1, 2, 3])\n");
}
#[test] fn next_basic() {
    compile_ok("items = [1, 2, 3]\nx = next(items)\n");
}

// ── Set methods ────────────────────────────────────────────

#[test] fn set_add() {
    compile_ok("s = {1, 2, 3}\ns.add(4)\n");
}
#[test] fn set_discard() {
    compile_ok("s = {1, 2, 3}\ns.discard(2)\n");
}
#[test] fn set_union() {
    compile_ok("a = {1, 2}\nb = {3, 4}\nc = a.union(b)\n");
}
#[test] fn set_intersection() {
    compile_ok("a = {1, 2, 3}\nb = {2, 3, 4}\nc = a.intersection(b)\n");
}
#[test] fn set_difference() {
    compile_ok("a = {1, 2, 3}\nb = {2, 3}\nc = a.difference(b)\n");
}

// ── @staticmethod / @classmethod ───────────────────────────

#[test] fn staticmethod_basic() {
    compile_ok("class Math:\n    @staticmethod\n    def add(a, b):\n        return a + b\nresult = Math.add(1, 2)\n");
}
#[test] fn staticmethod_no_self_param() {
    // Static method with no params at all
    compile_ok("class Config:\n    @staticmethod\n    def default_value():\n        return 42\nv = Config.default_value()\n");
}
#[test] fn staticmethod_with_instance_methods() {
    // Mix of static and instance methods
    compile_ok("class Counter:\n    def __init__(self, n):\n        self.n = n\n    def value(self):\n        return self.n\n    @staticmethod\n    def zero():\n        return Counter(0)\nc = Counter.zero()\nprint(c.value())\n");
}
#[test] fn classmethod_basic() {
    compile_ok("class Foo:\n    @classmethod\n    def create(cls):\n        return Foo()\n    def __init__(self):\n        pass\n");
}
#[test] fn classmethod_with_args() {
    compile_ok("class Point:\n    def __init__(self, x, y):\n        self.x = x\n        self.y = y\n    @classmethod\n    def origin(cls):\n        return Point(0, 0)\np = Point.origin()\n");
}

// ── f-string format specs ──────────────────────────────────

#[test] fn fstring_format_float() {
    compile_ok("x = 3.14159\ns = f\"{x:.2f}\"\n");
}
#[test] fn fstring_format_width() {
    compile_ok("name = \"hello\"\ns = f\"{name:>20}\"\n");
}
#[test] fn fstring_format_int() {
    compile_ok("n = 42\ns = f\"{n:04d}\"\n");
}
#[test] fn fstring_no_spec_still_works() {
    compile_ok("x = 10\ns = f\"value is {x}\"\n");
}

// ── dict.setdefault ────────────────────────────────────────

#[test] fn dict_setdefault_existing() {
    compile_ok("d = {'a': 1}\nv = d.setdefault('a', 99)\n");
}
#[test] fn dict_setdefault_missing() {
    compile_ok("d = {}\nv = d.setdefault('key', 42)\n");
}

// ── splitlines ─────────────────────────────────────────────

#[test] fn splitlines_basic() {
    compile_ok("text = \"hello\\nworld\\n\"\nlines = text.splitlines()\n");
}

// ── Cross-language compatibility ───────────────────────────

#[test] fn list_contains_method() {
    // Dart/C#/JS call .contains() — should work on Python lists
    compile_ok("x = [1, 2, 3]\nb = x.contains(2)\n");
}
#[test] fn list_includes_method() {
    // JS calls .includes() — should work on Python lists
    compile_ok("x = [1, 2, 3]\nb = x.includes(2)\n");
}
#[test] fn list_length_method() {
    // Dart calls .length — should work as method fallback
    compile_ok("x = [1, 2, 3]\nn = x.length()\n");
}
#[test] fn class_tostring_alias() {
    // Python __str__ should be callable as toString from JS/Dart
    compile_ok("class Dog:\n    def __init__(self, name):\n        self.name = name\n    def __str__(self):\n        return self.name\nd = Dog('Rex')\n");
}
#[test] fn class_contains_alias() {
    // Python __contains__ should be callable as contains/includes
    compile_ok("class Bag:\n    def __init__(self):\n        self.items = []\n    def __contains__(self, item):\n        return item in self.items\nb = Bag()\n");
}
