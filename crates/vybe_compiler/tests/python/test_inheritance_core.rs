use crate::helpers::run_python_one;

#[test]
fn inheritance_method_resolution_child() {
    assert_eq!(
        run_python_one("class B:\n def f(self):\n  return 'b'\nclass D(B):\n pass\nprint(D().f())\n"),
        "b"
    );
}

#[test]
fn inheritance_override_method() {
    assert_eq!(
        run_python_one("class B:\n def f(self):\n  return 1\nclass D(B):\n def f(self):\n  return 2\nprint(D().f())\n"),
        "2"
    );
}

#[test]
fn inheritance_super_call_parent() {
    assert_eq!(
        run_python_one("class B:\n def f(self):\n  return 1\nclass D(B):\n def f(self):\n  return super().f() + 1\nprint(D().f())\n"),
        "2"
    );
}

#[test]
fn inheritance_init_chain() {
    assert_eq!(
        run_python_one("class B:\n def __init__(self):\n  self.a = 1\nclass D(B):\n def __init__(self):\n  super().__init__()\n  self.b = 2\nd = D()\nprint(d.a, d.b)\n"),
        "1 2"
    );
}

#[test]
fn inheritance_class_attr_shared() {
    assert_eq!(
        run_python_one("class B:\n x = 1\nclass D(B):\n pass\nprint(D.x)\n"),
        "1"
    );
}

#[test]
fn inheritance_instance_attr_independent() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nb, d = B(), D()\nb.v = 1\nprint(hasattr(d, 'v'))\n"),
        "False"
    );
}

#[test]
fn inheritance_isinstance_child() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(isinstance(D(), B))\n"),
        "True"
    );
}

#[test]
fn inheritance_issubclass_direct() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(issubclass(D, B))\n"),
        "True"
    );
}

#[test]
fn inheritance_issubclass_same_class() {
    assert_eq!(
        run_python_one("class B:\n pass\nprint(issubclass(B, B))\n"),
        "True"
    );
}

#[test]
fn inheritance_mro_order() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint([c.__name__ for c in D.__mro__])\n"),
        "['D', 'B', 'object']"
    );
}

#[test]
fn inheritance_diamond_super_linearization() {
    assert_eq!(
        run_python_one(
            "class A:\n def f(self):\n  return 'A'\nclass B(A):\n def f(self):\n  return 'B' + super().f()\n\
             class C(A):\n def f(self):\n  return 'C' + super().f()\nclass D(B, C):\n def f(self):\n  return 'D' + super().f()\nprint(D().f())\n"
        ),
        "DBCA"
    );
}

#[test]
fn inheritance_call_parent_unbound_via_class() {
    assert_eq!(
        run_python_one("class B:\n def f(self):\n  return 3\nclass D(B):\n def f(self):\n  return B.f(self) + 1\nprint(D().f())\n"),
        "4"
    );
}

#[test]
fn inheritance_add_child_class_attr() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n y = 2\nprint(D.y)\n"),
        "2"
    );
}

#[test]
fn inheritance_shadow_parent_method_only_on_child() {
    assert_eq!(
        run_python_one("class B:\n def f(self):\n  return 'b'\nclass D(B):\n def f(self):\n  return 'd'\nb, d = B(), D()\nprint(b.f(), d.f())\n"),
        "b d"
    );
}

#[test]
fn inheritance_type_check_exact_child() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(type(D()) is D)\n"),
        "True"
    );
}

#[test]
fn inheritance_type_check_not_parent() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(type(D()) is B)\n"),
        "False"
    );
}

#[test]
fn inheritance_override_str() {
    assert_eq!(
        run_python_one("class B:\n def __str__(self):\n  return 'b'\nclass D(B):\n def __str__(self):\n  return 'd'\nprint(str(D()))\n"),
        "d"
    );
}

#[test]
fn inheritance_keep_parent_str() {
    assert_eq!(
        run_python_one("class B:\n def __str__(self):\n  return 'b'\nclass D(B):\n pass\nprint(str(D()))\n"),
        "b"
    );
}

#[test]
fn inheritance_super_in_property_getter() {
    assert_eq!(
        run_python_one("class B:\n def val(self):\n  return 1\nclass D(B):\n def val(self):\n  return super().val() + 1\nd = D()\nprint(d.val())\n"),
        "2"
    );
}

#[test]
fn inheritance_classmethod_inherited() {
    assert_eq!(
        run_python_one("class B:\n @classmethod\n def make(cls):\n  return cls()\nclass D(B):\n pass\nprint(isinstance(D.make(), D))\n"),
        "True"
    );
}

#[test]
fn inheritance_staticmethod_inherited() {
    assert_eq!(
        run_python_one("class B:\n @staticmethod\n def twice(x):\n  return x * 2\nclass D(B):\n pass\nprint(D.twice(3))\n"),
        "6"
    );
}

#[test]
fn inheritance_override_classmethod() {
    assert_eq!(
        run_python_one("class B:\n @classmethod\n def name(cls):\n  return 'B'\nclass D(B):\n @classmethod\n def name(cls):\n  return 'D'\nprint(D.name())\n"),
        "D"
    );
}

#[test]
fn inheritance_multiple_bases_methods() {
    assert_eq!(
        run_python_one("class A:\n def a(self):\n  return 1\nclass B:\n def b(self):\n  return 2\nclass C(A, B):\n pass\nc = C()\nprint(c.a(), c.b())\n"),
        "1 2"
    );
}

#[test]
fn inheritance_first_base_method_used() {
    assert_eq!(
        run_python_one("class A:\n def f(self):\n  return 'A'\nclass B:\n def f(self):\n  return 'B'\nclass C(A, B):\n pass\nprint(C().f())\n"),
        "A"
    );
}

#[test]
fn inheritance_second_base_after_super_chain() {
    assert_eq!(
        run_python_one("class A:\n def f(self):\n  return 'A'\nclass B:\n def f(self):\n  return 'B'\nclass C(B, A):\n pass\nprint(C().f())\n"),
        "B"
    );
}

#[test]
fn inheritance_setattr_on_child() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nd = D()\nd.x = 5\nprint(d.x)\n"),
        "5"
    );
}

#[test]
fn inheritance_getattr_fallback() {
    assert_eq!(
        run_python_one("class B:\n z = 9\nclass D(B):\n pass\nprint(D().z)\n"),
        "9"
    );
}

#[test]
fn inheritance_delattr_on_child_field() {
    assert_eq!(
        run_python_one("class B:\n def __init__(self):\n  self.x = 1\nclass D(B):\n pass\nd = D()\ndel d.x\nprint(hasattr(d, 'x'))\n"),
        "False"
    );
}

#[test]
fn inheritance_repr_includes_class_name() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint('D' in repr(D()))\n"),
        "True"
    );
}

#[test]
fn inheritance_len_from_parent() {
    assert_eq!(
        run_python_one("class B:\n def __len__(self):\n  return 4\nclass D(B):\n pass\nprint(len(D()))\n"),
        "4"
    );
}

#[test]
fn inheritance_iter_from_parent() {
    assert_eq!(
        run_python_one("class B:\n def __iter__(self):\n  return iter([1, 2])\nclass D(B):\n pass\nprint(list(D()))\n"),
        "[1, 2]"
    );
}

#[test]
fn inheritance_contains_from_parent() {
    assert_eq!(
        run_python_one("class B:\n def __contains__(self, item):\n  return item == 1\nclass D(B):\n pass\nprint(1 in D())\n"),
        "True"
    );
}

#[test]
fn inheritance_call_from_parent() {
    assert_eq!(
        run_python_one("class B:\n def __call__(self, x):\n  return x + 1\nclass D(B):\n pass\nprint(D()(3))\n"),
        "4"
    );
}

#[test]
fn inheritance_override_init_only_child() {
    assert_eq!(
        run_python_one("class B:\n def __init__(self):\n  self.flag = 'b'\nclass D(B):\n def __init__(self):\n  self.flag = 'd'\nprint(D().flag)\n"),
        "d"
    );
}

#[test]
fn inheritance_super_without_override_in_middle() {
    assert_eq!(
        run_python_one("class A:\n def f(self):\n  return 1\nclass B(A):\n pass\nclass C(B):\n def f(self):\n  return super().f() + 1\nprint(C().f())\n"),
        "2"
    );
}

#[test]
fn inheritance_object_in_mro() {
    assert_eq!(
        run_python_one("class B:\n pass\nprint(B.__mro__[-1].__name__)\n"),
        "object"
    );
}

#[test]
fn inheritance_child_adds_new_method() {
    assert_eq!(
        run_python_one("class B:\n def b(self):\n  return 1\nclass D(B):\n def d(self):\n  return 2\nd = D()\nprint(d.b(), d.d())\n"),
        "1 2"
    );
}

#[test]
fn inheritance_parent_keeps_old_behavior() {
    assert_eq!(
        run_python_one("class B:\n def f(self):\n  return 1\nclass D(B):\n def f(self):\n  return 2\nprint(B().f())\n"),
        "1"
    );
}

#[test]
fn inheritance_bound_super_in_method() {
    assert_eq!(
        run_python_one("class B:\n def values(self):\n  return [1]\nclass D(B):\n def values(self):\n  return super().values() + [2]\nprint(D().values())\n"),
        "[1, 2]"
    );
}

#[test]
fn inheritance_class_name_on_child() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(D.__name__)\n"),
        "D"
    );
}

#[test]
fn inheritance_base_list_on_child() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(D.__bases__[0].__name__)\n"),
        "B"
    );
}

#[test]
fn inheritance_polymorphic_call() {
    assert_eq!(
        run_python_one("class B:\n def speak(self):\n  return 'b'\nclass D(B):\n def speak(self):\n  return 'd'\ndef say(x):\n return x.speak()\nprint(say(B()), say(D()))\n"),
        "b d"
    );
}

#[test]
fn inheritance_super_passes_args() {
    assert_eq!(
        run_python_one("class B:\n def __init__(self, n):\n  self.n = n\nclass D(B):\n def __init__(self, n):\n  super().__init__(n)\nprint(D(5).n)\n"),
        "5"
    );
}

#[test]
fn inheritance_three_level_chain() {
    assert_eq!(
        run_python_one("class A:\n def f(self):\n  return 'A'\nclass B(A):\n def f(self):\n  return super().f() + 'B'\nclass C(B):\n def f(self):\n  return super().f() + 'C'\nprint(C().f())\n"),
        "ABC"
    );
}

#[test]
fn inheritance_instance_dict_separate() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nb, d = B(), D()\nb.x = 1\nd.x = 2\nprint(b.x, d.x)\n"),
        "1 2"
    );
}
