use crate::helpers::run_python_one;

#[test]
fn class_instance_attribute() {
    assert_eq!(
        run_python_one("class A:\n pass\na = A()\na.x = 5\nprint(a.x)\n"),
        "5"
    );
}

#[test]
fn class_method_returns_self_field() {
    assert_eq!(
        run_python_one(
            "class Counter:\n def __init__(self):\n  self.n = 0\n c = Counter()\nprint(c.n)\n"
        ),
        "0"
    );
}

#[test]
fn class_init_sets_fields() {
    assert_eq!(
        run_python_one(
            "class Point:\n def __init__(self, x, y):\n  self.x = x\n  self.y = y\np = Point(1, 2)\nprint(p.x, p.y)\n"
        ),
        "1 2"
    );
}

#[test]
fn class_method_on_instance() {
    assert_eq!(
        run_python_one(
            "class Greeter:\n def hi(self):\n  return 'hi'\ng = Greeter()\nprint(g.hi())\n"
        ),
        "hi"
    );
}

#[test]
fn class_method_mutates_state() {
    assert_eq!(
        run_python_one(
            "class Acc:\n def __init__(self):\n  self.v = 0\n def inc(self):\n  self.v += 1\na = Acc()\na.inc()\nprint(a.v)\n"
        ),
        "1"
    );
}

#[test]
fn class_class_attribute_shared() {
    assert_eq!(run_python_one("class C:\n tag = 'x'\nprint(C.tag)\n"), "x");
}

#[test]
fn class_instance_sees_class_attr() {
    assert_eq!(
        run_python_one("class C:\n tag = 7\nc = C()\nprint(c.tag)\n"),
        "7"
    );
}

#[test]
fn class_instance_attr_shadows_class() {
    assert_eq!(
        run_python_one("class C:\n x = 1\nc = C()\nc.x = 2\nprint(c.x, C.x)\n"),
        "2 1"
    );
}

#[test]
fn class_repr_default() {
    assert_eq!(
        run_python_one("class A:\n pass\nprint(A())\n").contains("A"),
        true
    );
}

#[test]
fn class_custom_str() {
    assert_eq!(
        run_python_one("class A:\n def __str__(self):\n  return 'custom'\nprint(str(A()))\n"),
        "custom"
    );
}

#[test]
fn class_equality_default_by_identity() {
    assert_eq!(
        run_python_one("class A:\n pass\na = A()\nb = A()\nprint(a == b)\n"),
        "False"
    );
}

#[test]
fn class_custom_eq() {
    assert_eq!(
        run_python_one(
            "class V:\n def __init__(self, n):\n  self.n = n\n def __eq__(self, o):\n  return self.n == o.n\nprint(V(1) == V(1))\n"
        ),
        "True"
    );
}

#[test]
fn class_property_style_getter() {
    assert_eq!(
        run_python_one(
            "class R:\n def __init__(self, x):\n  self._x = x\n def get(self):\n  return self._x\nr = R(3)\nprint(r.get())\n"
        ),
        "3"
    );
}

#[test]
fn class_static_method_call() {
    assert_eq!(
        run_python_one(
            "class U:\n @staticmethod\n def twice(x):\n  return x * 2\nprint(U.twice(4))\n"
        ),
        "8"
    );
}

#[test]
fn class_classmethod_factory() {
    assert_eq!(
        run_python_one(
            "class A:\n @classmethod\n def make(cls):\n  return cls()\nprint(isinstance(A.make(), A))\n"
        ),
        "True"
    );
}

#[test]
fn class_inheritance_method_override() {
    assert_eq!(
        run_python_one(
            "class B:\n def f(self):\n  return 1\nclass D(B):\n def f(self):\n  return 2\nprint(D().f())\n"
        ),
        "2"
    );
}

#[test]
fn class_inheritance_super_call() {
    assert_eq!(
        run_python_one(
            "class B:\n def f(self):\n  return 1\nclass D(B):\n def f(self):\n  return super().f() + 1\nprint(D().f())\n"
        ),
        "2"
    );
}

#[test]
fn class_private_name_mangling_behavior() {
    assert_eq!(
        run_python_one("class C:\n def __init__(self):\n  self.__x = 9\nc = C()\nprint(c._C__x)\n"),
        "9"
    );
}

#[test]
fn class_slots_not_used_default_dict() {
    assert_eq!(
        run_python_one("class C:\n pass\nc = C()\nc.a = 1\nprint(c.a)\n"),
        "1"
    );
}

#[test]
fn class_method_with_default_arg() {
    assert_eq!(
        run_python_one("class F:\n def add(self, x, y=1):\n  return x + y\nprint(F().add(2))\n"),
        "3"
    );
}

#[test]
fn class_iter_simple() {
    assert_eq!(
        run_python_one(
            "class R:\n def __init__(self, xs):\n  self.xs = xs\n def __iter__(self):\n  return iter(self.xs)\nprint(list(R([1, 2])))\n"
        ),
        "[1, 2]"
    );
}

#[test]
fn class_len_dunder() {
    assert_eq!(
        run_python_one("class B:\n def __len__(self):\n  return 3\nprint(len(B()))\n"),
        "3"
    );
}

#[test]
fn class_bool_dunder() {
    assert_eq!(
        run_python_one("class B:\n def __bool__(self):\n  return False\nprint(bool(B()))\n"),
        "False"
    );
}

#[test]
fn class_contains_dunder() {
    assert_eq!(
        run_python_one(
            "class B:\n def __contains__(self, item):\n  return item == 1\nprint(1 in B())\n"
        ),
        "True"
    );
}

#[test]
fn class_getitem_dunder() {
    assert_eq!(
        run_python_one("class B:\n def __getitem__(self, i):\n  return i * 2\nprint(B()[3])\n"),
        "6"
    );
}

#[test]
fn class_call_dunder() {
    assert_eq!(
        run_python_one("class F:\n def __call__(self, x):\n  return x + 1\nprint(F()(4))\n"),
        "5"
    );
}

#[test]
fn class_multiple_instances_independent() {
    assert_eq!(
        run_python_one(
            "class C:\n def __init__(self):\n  self.v = []\na, b = C(), C()\na.v.append(1)\nprint(b.v)\n"
        ),
        "[]"
    );
}

#[test]
fn class_method_binding() {
    assert_eq!(
        run_python_one("class C:\n def f(self):\n  return 1\nc = C()\nm = c.f\nprint(m())\n"),
        "1"
    );
}

#[test]
fn class_attribute_on_class_from_instance() {
    assert_eq!(
        run_python_one("class C:\n n = 0\nc = C()\nC.n = 5\nprint(c.n)\n"),
        "5"
    );
}

#[test]
fn class_nested_class() {
    assert_eq!(
        run_python_one("class Outer:\n class Inner:\n  v = 1\nprint(Outer.Inner.v)\n"),
        "1"
    );
}

#[test]
fn class_instance_dict_dynamic() {
    assert_eq!(
        run_python_one(
            "class C:\n pass\nc = C()\nc.x = 1\nc.y = 2\nprint(sorted(c.__dict__.keys()))\n"
        ),
        "['x', 'y']"
    );
}

#[test]
fn class_isinstance_check() {
    assert_eq!(
        run_python_one("class A:\n pass\nprint(isinstance(A(), A))\n"),
        "True"
    );
}

#[test]
fn class_issubclass_check() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(issubclass(D, B))\n"),
        "True"
    );
}

#[test]
fn class_method_returns_new_instance() {
    assert_eq!(
        run_python_one(
            "class B:\n def copy(self):\n  return B()\nprint(isinstance(B().copy(), B))\n"
        ),
        "True"
    );
}

#[test]
fn class_del_method_optional() {
    assert_eq!(
        run_python_one("class C:\n def __init__(self):\n  self.ok = True\nc = C()\nprint(c.ok)\n"),
        "True"
    );
}

#[test]
fn class_repr_with_fields() {
    assert_eq!(
        run_python_one(
            "class P:\n def __init__(self, x):\n  self.x = x\n def __repr__(self):\n  return f'P({self.x})'\nprint(repr(P(2)))\n"
        ),
        "P(2)"
    );
}

#[test]
fn class_comparison_not_implemented_false() {
    assert_eq!(
        run_python_one("class A:\n pass\nclass B:\n pass\nprint(A() < B())\n"),
        "False"
    );
}

#[test]
fn class_richcompare_lt() {
    assert_eq!(
        run_python_one(
            "class N:\n def __init__(self, v):\n  self.v = v\n def __lt__(self, o):\n  return self.v < o.v\nprint(N(1) < N(2))\n"
        ),
        "True"
    );
}

#[test]
fn class_hash_disabled_by_default() {
    assert_eq!(
        run_python_one(
            "class C:\n def __init__(self):\n  self.x = 1\ntry:\n hash(C())\nexcept TypeError:\n print('no')\n"
        ),
        "no"
    );
}

#[test]
fn class_explicit_hash() {
    assert_eq!(
        run_python_one(
            "class C:\n def __init__(self, x):\n  self.x = x\n def __hash__(self):\n  return hash(self.x)\nprint(hash(C(1)) == hash(C(1)))\n"
        ),
        "True"
    );
}

#[test]
fn class_dataclass_style_manual() {
    assert_eq!(
        run_python_one(
            "class P:\n def __init__(self, x, y):\n  self.x = x\n  self.y = y\np = P(1, 2)\nprint(p.x + p.y)\n"
        ),
        "3"
    );
}

#[test]
fn class_method_chaining() {
    assert_eq!(
        run_python_one(
            "class B:\n def __init__(self):\n  self.n = 0\n def inc(self):\n  self.n += 1\n  return self\nprint(B().inc().inc().n)\n"
        ),
        "2"
    );
}

#[test]
fn class_super_init_chain() {
    assert_eq!(
        run_python_one(
            "class B:\n def __init__(self):\n  self.a = 1\nclass D(B):\n def __init__(self):\n  super().__init__()\n  self.b = 2\nd = D()\nprint(d.a, d.b)\n"
        ),
        "1 2"
    );
}

#[test]
fn class_type_name() {
    assert_eq!(
        run_python_one("class MyCls:\n pass\nprint(MyCls.__name__)\n"),
        "MyCls"
    );
}

#[test]
fn class_mro_length() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(len(D.__mro__))\n"),
        "3"
    );
}

#[test]
fn class_instance_method_type() {
    assert_eq!(
        run_python_one("class C:\n def f(self):\n  pass\nprint(type(C().f).__name__)\n"),
        "method"
    );
}
