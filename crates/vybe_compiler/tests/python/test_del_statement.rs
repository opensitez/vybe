use crate::helpers::run_python_one;

#[test]
fn del_list_item() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\ndel a[1]\nprint(a)\n"),
        "[1, 3]"
    );
}

#[test]
fn del_list_slice() {
    assert_eq!(
        run_python_one("a = [1, 2, 3, 4]\ndel a[1:3]\nprint(a)\n"),
        "[1, 4]"
    );
}

#[test]
fn del_dict_key() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2}\ndel d['a']\nprint(d)\n"),
        "{'b': 2}"
    );
}

#[test]
fn del_variable() {
    assert_eq!(
        run_python_one("x = 1\ndel x\ntry:\n print(x)\nexcept NameError:\n print('gone')\n"),
        "gone"
    );
}

#[test]
fn del_attr_on_object() {
    assert_eq!(
        run_python_one("class C:\n pass\nc = C()\nc.x = 1\ndel c.x\nprint(hasattr(c, 'x'))\n"),
        "False"
    );
}

#[test]
fn del_slice_step() {
    assert_eq!(
        run_python_one("a = [0, 1, 2, 3, 4]\ndel a[::2]\nprint(a)\n"),
        "[1, 3]"
    );
}

#[test]
fn del_empty_slice_noop() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\ndel a[1:1]\nprint(a)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn del_list_index_error() {
    assert_eq!(
        run_python_one("try:\n a = []\n del a[0]\nexcept IndexError:\n print('idx')\n"),
        "idx"
    );
}

#[test]
fn del_dict_key_error() {
    assert_eq!(
        run_python_one("try:\n del {}['x']\nexcept KeyError:\n print('key')\n"),
        "key"
    );
}

#[test]
fn del_name_error_undefined() {
    assert_eq!(
        run_python_one("try:\n del not_defined\nexcept NameError:\n print('name')\n"),
        "name"
    );
}

#[test]
fn del_multiple_targets_unbind() {
    assert_eq!(
        run_python_one("a = b = 1\ndel a, b\ntry:\n print(a)\nexcept NameError:\n print('ok')\n"),
        "ok"
    );
}

#[test]
fn del_nested_list_item() {
    assert_eq!(
        run_python_one("m = [[1, 2], [3]]\ndel m[0][1]\nprint(m)\n"),
        "[[1], [3]]"
    );
}

#[test]
fn del_set_variable() {
    assert_eq!(run_python_one("s = {1, 2}\ndel s\nprint('ok')\n"), "ok");
}

#[test]
fn del_set_item_not_supported() {
    assert_eq!(
        run_python_one("try:\n s = {1, 2}\n del s[0]\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}

#[test]
fn del_class_attribute_from_instance_dict_only() {
    assert_eq!(
        run_python_one("class C:\n x = 1\nc = C()\nc.x = 2\ndel c.x\nprint(c.x)\n"),
        "1"
    );
}

#[test]
fn del_class_attribute_from_class() {
    assert_eq!(
        run_python_one(
            "class C:\n x = 1\ndel C.x\ntry:\n print(C.x)\nexcept AttributeError:\n print('attr')\n"
        ),
        "attr"
    );
}

#[test]
fn del_subclass_attr_shadow() {
    assert_eq!(
        run_python_one("class B:\n x = 1\nclass D(B):\n x = 2\ndel D.x\nprint(D.x)\n"),
        "1"
    );
}

#[test]
fn del_after_pop_equivalent() {
    assert_eq!(run_python_one("a = [1, 2]\na.pop()\nprint(a)\n"), "[1]");
}

#[test]
fn del_tuple_index_not_allowed() {
    assert_eq!(
        run_python_one("try:\n t = (1, 2)\n del t[0]\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}

#[test]
fn del_string_item_not_allowed() {
    assert_eq!(
        run_python_one("try:\n s = 'ab'\n del s[0]\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}

#[test]
fn del_list_all_via_slice() {
    assert_eq!(run_python_one("a = [1, 2, 3]\ndel a[:]\nprint(a)\n"), "[]");
}

#[test]
fn del_dict_clear_vs_del() {
    assert_eq!(run_python_one("d = {'a': 1}\ndel d['a']\nprint(d)\n"), "{}");
}

#[test]
fn del_in_function_local() {
    assert_eq!(
        run_python_one("def f():\n x = 1\n del x\n return 'ok'\nprint(f())\n"),
        "ok"
    );
}

#[test]
fn del_global_name() {
    assert_eq!(
        run_python_one(
            "g = 1\ndef f():\n global g\n del g\nf()\ntry:\n print(g)\nexcept NameError:\n print('gone')\n"
        ),
        "gone"
    );
}

#[test]
fn del_nonlocal_name() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 1\n def inner():\n  nonlocal x\n  del x\n inner()\n try:\n  return x\n except NameError:\n  return 'gone'\nprint(outer())\n"
        ),
        "gone"
    );
}

#[test]
fn del_list_negative_index() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\ndel a[-1]\nprint(a)\n"),
        "[1, 2]"
    );
}

#[test]
fn del_dict_item_in_loop() {
    assert_eq!(
        run_python_one(
            "d = {'a': 1, 'b': 2}\nfor k in list(d):\n if k == 'a':\n  del d[k]\nprint(d)\n"
        ),
        "{'b': 2}"
    );
}

#[test]
fn del_attr_error_on_missing() {
    assert_eq!(
        run_python_one(
            "class C:\n pass\ntry:\n del C.missing\nexcept AttributeError:\n print('attr')\n"
        ),
        "attr"
    );
}

#[test]
fn del_package_submodule_style_attr() {
    assert_eq!(
        run_python_one(
            "class M:\n value = 1\ndel M.value\ntry:\n print(M.value)\nexcept AttributeError:\n print('ok')\n"
        ),
        "ok"
    );
}

#[test]
fn del_slice_assignment_then_del() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\na[1] = 9\ndel a[1]\nprint(a)\n"),
        "[1, 3]"
    );
}

#[test]
fn del_comprehension_temp_not_applicable() {
    assert_eq!(
        run_python_one("a = [x for x in range(3)]\ndel a[0]\nprint(a)\n"),
        "[1, 2]"
    );
}

#[test]
fn del_two_step_rebind() {
    assert_eq!(run_python_one("a = [1]\nb = a\ndel a\nprint(b)\n"), "[1]");
}

#[test]
fn del_property_custom_deleter() {
    assert_eq!(
        run_python_one(
            "class C:\n def __init__(self):\n  self._x = 1\n @property\n def x(self):\n  return self._x\n @x.deleter\n def x(self):\n  del self._x\nc = C()\ndel c.x\nprint(hasattr(c, '_x'))\n"
        ),
        "False"
    );
}

#[test]
fn del_item_from_custom_mapping() {
    assert_eq!(
        run_python_one("class M(dict):\n pass\nm = M({'a': 1})\ndel m['a']\nprint(m)\n"),
        "{}"
    );
}

#[test]
fn del_last_reference_allows_gc() {
    assert_eq!(run_python_one("a = [1]\ndel a\nprint('ok')\n"), "ok");
}

#[test]
fn del_key_from_kwargs_copy() {
    assert_eq!(
        run_python_one("def f(**kw):\n del kw['a']\n return kw\nprint(f(a=1, b=2))\n"),
        "{'b': 2}"
    );
}

#[test]
fn del_from_list_while_iterating_copy() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\nfor x in list(a):\n if x == 2:\n  a.remove(x)\nprint(a)\n"),
        "[1, 3]"
    );
}

#[test]
fn del_extended_slice() {
    assert_eq!(
        run_python_one("a = list(range(6))\ndel a[1:5:2]\nprint(a)\n"),
        "[0, 2, 4, 5]"
    );
}

#[test]
fn del_rebind_after_del_name() {
    assert_eq!(run_python_one("x = 1\ndel x\nx = 2\nprint(x)\n"), "2");
}

#[test]
fn del_empty_dict_key_error() {
    assert_eq!(
        run_python_one("try:\n del {}[0]\nexcept KeyError:\n print('key')\n"),
        "key"
    );
}

#[test]
fn del_bytearray_item() {
    assert_eq!(
        run_python_one(
            "try:\n ba = bytearray(b'ab')\n del ba[0]\n print(ba)\nexcept:\n print('err')\n"
        ),
        "bytearray(b'b')"
    );
}

#[test]
fn del_module_level_list_item() {
    assert_eq!(
        run_python_one("items = [1, 2, 3]\ndel items[1]\nprint(items)\n"),
        "[1, 3]"
    );
}

#[test]
fn del_chained_attr() {
    assert_eq!(
        run_python_one(
            "class A:\n def __init__(self):\n  self.b = {'k': 1}\na = A()\ndel a.b['k']\nprint(a.b)\n"
        ),
        "{}"
    );
}

#[test]
fn del_overwrites_then_del() {
    assert_eq!(
        run_python_one("d = {'a': 1}\nd['a'] = 2\ndel d['a']\nprint(d)\n"),
        "{}"
    );
}

#[test]
fn del_function_local_after_return_unreachable() {
    assert_eq!(
        run_python_one("def f():\n x = 1\n return x\ndel f\nprint('ok')\n"),
        "ok"
    );
}
