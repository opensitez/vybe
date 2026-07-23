use crate::helpers::{run_print, run_python_one};

#[test]
fn is_same_object() {
    assert_eq!(run_print("[] is []"), "False");
}

#[test]
fn is_not_different_objects() {
    assert_eq!(run_print("[1] is not [1]"), "True");
}

#[test]
fn is_same_literal_interned_small_int() {
    assert_eq!(run_print("256 is 256"), "True");
}

#[test]
fn is_small_int_cached() {
    assert_eq!(run_print("100 is 100"), "True");
}

#[test]
fn is_none_singleton() {
    assert_eq!(run_print("None is None"), "True");
}

#[test]
fn is_true_singleton() {
    assert_eq!(run_print("True is True"), "True");
}

#[test]
fn is_false_singleton() {
    assert_eq!(run_print("False is False"), "True");
}

#[test]
fn is_variable_self() {
    assert_eq!(run_python_one("x = []\nprint(x is x)\n"), "True");
}

#[test]
fn is_not_with_variables() {
    assert_eq!(
        run_python_one("a = [1]\nb = [1]\nprint(a is not b)\n"),
        "True"
    );
}

#[test]
fn is_equal_but_not_identical() {
    assert_eq!(
        run_python_one("a = [1]\nb = [1]\nprint(a == b, a is b)\n"),
        "True False"
    );
}

#[test]
fn is_function_same_reference() {
    assert_eq!(
        run_python_one("def f():\n pass\ng = f\nprint(f is g)\n"),
        "True"
    );
}

#[test]
fn is_class_same_type() {
    assert_eq!(
        run_python_one("class A:\n pass\nprint(A() is A())\n"),
        "False"
    );
}

#[test]
fn is_string_interned_literal() {
    assert_eq!(run_print("'hello' is 'hello'"), "True");
}

#[test]
fn is_not_none_check() {
    assert_eq!(run_python_one("x = None\nprint(x is not None)\n"), "False");
}

#[test]
fn is_not_guard_pattern() {
    assert_eq!(
        run_python_one("x = []\nprint('ok' if x is not None else 'no')\n"),
        "ok"
    );
}

#[test]
fn is_tuple_not_same_each_time() {
    assert_eq!(run_print("(1,) is (1,)"), "False");
}

#[test]
fn is_dict_not_same() {
    assert_eq!(run_print("{} is {}"), "False");
}

#[test]
fn is_set_not_same() {
    assert_eq!(run_print("{1} is {1}"), "False");
}

#[test]
fn is_bound_method_same() {
    assert_eq!(
        run_python_one("class C:\n def f(self):\n  pass\nc = C()\nprint(c.f is c.f)\n"),
        "True"
    );
}

#[test]
fn is_not_bound_method_different_instances() {
    assert_eq!(
        run_python_one("class C:\n def f(self):\n  pass\nprint(C().f is C().f)\n"),
        "False"
    );
}

#[test]
fn is_aliased_list() {
    assert_eq!(run_python_one("a = [1]\nb = a\nprint(a is b)\n"), "True");
}

#[test]
fn is_after_assignment() {
    assert_eq!(
        run_python_one("a = object()\nb = a\nprint(a is b)\n"),
        "True"
    );
}

#[test]
fn is_not_after_rebind() {
    assert_eq!(
        run_python_one("a = []\nb = a\na = []\nprint(b is a)\n"),
        "False"
    );
}

#[test]
fn is_zero_int() {
    assert_eq!(run_print("0 is 0"), "True");
}

#[test]
fn is_negative_one_cached() {
    assert_eq!(run_print("-5 is -5"), "True");
}

#[test]
fn is_float_not_cached() {
    assert_eq!(run_print("1.0 is 1.0"), "True");
}

#[test]
fn is_empty_string_interned() {
    assert_eq!(run_print("'' is ''"), "True");
}

#[test]
fn is_in_if_identity_check() {
    assert_eq!(
        run_python_one(
            "sentinel = object()\nvalue = sentinel\nprint('same' if value is sentinel else 'diff')\n"
        ),
        "same"
    );
}

#[test]
fn is_not_in_filter() {
    assert_eq!(
        run_python_one(
            "a = [1, 2]\nb = a\npairs = [(a, b), ([1], [1])]\nprint(sum(1 for x, y in pairs if x is y))\n"
        ),
        "1"
    );
}

#[test]
fn is_class_object_singleton() {
    assert_eq!(run_python_one("class A:\n pass\nprint(A is A)\n"), "True");
}

#[test]
fn is_module_level_same_name_rebind() {
    assert_eq!(
        run_python_one("x = []\ny = x\nx = []\nprint(y is x)\n"),
        "False"
    );
}

#[test]
fn is_compare_with_eq_diff_types() {
    assert_eq!(run_print("1 is 1.0"), "False");
}

#[test]
fn is_not_compare_with_eq_diff_types() {
    assert_eq!(run_print("1 is not 1.0"), "True");
}

#[test]
fn is_bytes_literal() {
    assert_eq!(run_print("b'a' is b'a'"), "True");
}

#[test]
fn is_frozenset_not_same() {
    assert_eq!(run_print("frozenset({1}) is frozenset({1})"), "False");
}

#[test]
fn is_lambda_not_same() {
    assert_eq!(run_print("(lambda: 0) is (lambda: 0)"), "False");
}

#[test]
fn is_generator_expression_not_same() {
    assert_eq!(
        run_python_one("print((x for x in range(1)) is (x for x in range(1)))\n"),
        "False"
    );
}

#[test]
fn is_slice_object() {
    assert_eq!(run_print("slice(1) is slice(1)"), "False");
}

#[test]
fn is_ellipsis_singleton() {
    assert_eq!(run_print("... is ..."), "True");
}

#[test]
fn is_not_ellipsis() {
    assert_eq!(run_print("None is not ..."), "True");
}

#[test]
fn is_chained_with_and() {
    assert_eq!(
        run_python_one("a = b = []\nprint(a is b and b is a)\n"),
        "True"
    );
}

#[test]
fn is_function_default_arg_mutable_trap() {
    assert_eq!(
        run_python_one("def f(a=[]):\n return a\ng = f()\nprint(g is f())\n"),
        "True"
    );
}

#[test]
fn is_cell_in_closure() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = []\n def inner():\n  return x\n a = inner()\n b = inner()\n print(a is b)\nprint(outer())\n"
        ),
        // `print(a is b)` → True (same captured cell), then `print(outer())`
        // prints outer's implicit None return.
        "True\nNone"
    );
}

#[test]
fn is_not_for_optional_none() {
    assert_eq!(
        run_python_one("def f(x=None):\n return x is not None\nprint(f(), f(1))\n"),
        "False True"
    );
}

#[test]
fn is_type_of_object() {
    assert_eq!(run_print("type(1) is int"), "True");
}

#[test]
fn is_subclass_not_is_instance() {
    assert_eq!(
        run_python_one("class B:\n pass\nclass D(B):\n pass\nprint(D is B, isinstance(D(), B))\n"),
        "False True"
    );
}
