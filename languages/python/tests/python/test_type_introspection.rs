use crate::helpers::{run_print, run_python_one};

#[test]
fn type_int_name() {
    assert_eq!(run_print("type(1).__name__"), "int");
}

#[test]
fn type_str_name() {
    assert_eq!(run_print("type('a').__name__"), "str");
}

#[test]
fn type_list_name() {
    assert_eq!(run_print("type([]).__name__"), "list");
}

#[test]
fn type_dict_name() {
    assert_eq!(run_print("type({}).__name__"), "dict");
}

#[test]
fn isinstance_int() {
    assert_eq!(run_print("isinstance(1, int)"), "True");
}

#[test]
fn isinstance_bool_subclass_int() {
    assert_eq!(run_print("isinstance(True, int)"), "True");
}

#[test]
fn isinstance_str_not_int() {
    assert_eq!(run_print("isinstance('1', int)"), "False");
}

#[test]
fn isinstance_tuple_of_types() {
    assert_eq!(run_print("isinstance(1, (str, int))"), "True");
}

#[test]
fn issubclass_bool_int() {
    assert_eq!(run_print("issubclass(bool, int)"), "True");
}

#[test]
fn issubclass_int_bool_false() {
    assert_eq!(run_print("issubclass(int, bool)"), "False");
}

#[test]
fn issubclass_list_object() {
    assert_eq!(run_print("issubclass(list, object)"), "True");
}

#[test]
fn callable_function() {
    assert_eq!(run_print("callable(print)"), "True");
}

#[test]
fn callable_int_not() {
    assert_eq!(run_print("callable(1)"), "False");
}

#[test]
fn callable_lambda() {
    assert_eq!(run_print("callable(lambda: 0)"), "True");
}

#[test]
fn callable_class() {
    assert_eq!(run_print("callable(int)"), "True");
}

#[test]
fn callable_instance_with_call() {
    assert_eq!(
        run_python_one("class C:\n def __call__(self):\n  pass\nprint(callable(C()))\n"),
        "True"
    );
}

#[test]
fn hasattr_on_object() {
    assert_eq!(run_print("hasattr([], 'append')"), "True");
}

#[test]
fn hasattr_missing() {
    assert_eq!(run_print("hasattr(1, 'append')"), "False");
}

#[test]
fn getattr_existing() {
    assert_eq!(run_print("getattr([], 'append') is list.append"), "False");
}

#[test]
fn getattr_default() {
    assert_eq!(run_print("getattr(1, 'missing', 9)"), "9");
}

#[test]
fn setattr_dynamic() {
    assert_eq!(
        run_python_one("class C:\n pass\nc = C()\nsetattr(c, 'x', 1)\nprint(c.x)\n"),
        "1"
    );
}

#[test]
fn delattr_removes() {
    assert_eq!(
        run_python_one("class C:\n x = 1\nc = C()\ndelattr(c, 'x')\nprint(hasattr(c, 'x'))\n"),
        "False"
    );
}

#[test]
fn dir_builtins_contains_len() {
    assert_eq!(
        run_python_one("print('len' in dir(__builtins__))\n"),
        "True"
    );
}

#[test]
fn dir_object_lists_attrs() {
    assert_eq!(
        run_python_one("class C:\n x = 1\nprint('x' in dir(C()))\n"),
        "True"
    );
}

#[test]
fn vars_on_object_dict() {
    assert_eq!(
        run_python_one("class C:\n pass\nc = C()\nc.a = 1\nprint(vars(c))\n"),
        "{'a': 1}"
    );
}

#[test]
fn id_same_object_equal() {
    assert_eq!(run_python_one("x = []\nprint(id(x) == id(x))\n"), "True");
}

#[test]
fn id_diff_objects() {
    assert_eq!(run_python_one("print(id([]) == id([]))\n"), "False");
}

#[test]
fn isinstance_custom_class() {
    assert_eq!(
        run_python_one("class A:\n pass\nprint(isinstance(A(), A))\n"),
        "True"
    );
}

#[test]
fn issubclass_custom_hierarchy() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(issubclass(D, B))\n"),
        "True"
    );
}

#[test]
fn type_of_class_is_type() {
    assert_eq!(
        run_python_one("class A:\n pass\nprint(type(A) is type)\n"),
        "True"
    );
}

#[test]
fn type_of_instance_is_class() {
    assert_eq!(
        run_python_one("class A:\n pass\nprint(type(A()).__name__)\n"),
        "A"
    );
}

#[test]
fn isinstance_none_type() {
    assert_eq!(run_print("isinstance(None, type(None))"), "True");
}

#[test]
fn isinstance_exception_hierarchy() {
    assert_eq!(run_print("isinstance(ValueError(), Exception)"), "True");
}

#[test]
fn issubclass_value_error_exception() {
    assert_eq!(run_print("issubclass(ValueError, Exception)"), "True");
}

#[test]
fn callable_bound_method() {
    assert_eq!(
        run_python_one("class C:\n def f(self):\n  pass\nprint(callable(C().f))\n"),
        "True"
    );
}

#[test]
fn getattr_class_attr() {
    assert_eq!(run_print("getattr(str, 'upper') is str.upper"), "True");
}

#[test]
fn hasattr_dunder_len() {
    assert_eq!(run_print("hasattr('', '__len__')"), "True");
}

#[test]
fn type_compare_with_is() {
    assert_eq!(run_print("type(1) is int"), "True");
}

#[test]
fn isinstance_tuple_empty_false() {
    assert_eq!(run_print("isinstance(1, ())"), "False");
}

#[test]
fn issubclass_same_class() {
    assert_eq!(run_print("issubclass(int, int)"), "True");
}

#[test]
fn isinstance_list_subclass() {
    assert_eq!(run_print("isinstance([], list)"), "True");
}

#[test]
fn isinstance_user_subclass() {
    assert_eq!(
        run_python_one("class L(list):\n pass\nprint(isinstance(L(), list))\n"),
        "True"
    );
}

#[test]
fn type_name_on_function() {
    assert_eq!(
        run_python_one("def f():\n pass\nprint(type(f).__name__)\n"),
        "function"
    );
}

#[test]
fn type_name_on_lambda() {
    assert_eq!(run_print("type(lambda: 0).__name__"), "function");
}

#[test]
fn callable_class_instance_without_call_false() {
    assert_eq!(
        run_python_one("class C:\n pass\nprint(callable(C()))\n"),
        "False"
    );
}

#[test]
fn getattr_descriptor_on_class() {
    assert_eq!(run_print("getattr(int, 'real') is int.real"), "True");
}
