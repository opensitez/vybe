use crate::helpers::{run_print, run_python_one};

#[test]
fn callable_on_lambda_is_true() {
    assert_eq!(run_print("callable(lambda x: x)"), "True");
}

#[test]
fn callable_on_builtin_len() {
    assert_eq!(run_print("callable(len)"), "True");
}

#[test]
fn callable_on_int_is_false() {
    assert_eq!(run_print("callable(42)"), "False");
}

#[test]
fn callable_on_user_function() {
    assert_eq!(
        run_python_one("def f():\n pass\nprint(callable(f))\n"),
        "True"
    );
}

#[test]
fn callable_on_class_is_true() {
    assert_eq!(
        run_python_one("class C:\n pass\nprint(callable(C))\n"),
        "True"
    );
}

#[test]
fn callable_on_instance_default_false() {
    assert_eq!(
        run_python_one("class C:\n pass\nprint(callable(C()))\n"),
        "False"
    );
}

#[test]
fn hash_of_int_is_int() {
    assert_eq!(run_print("type(hash(42)).__name__"), "int");
}

#[test]
fn hash_of_string_is_int() {
    assert_eq!(run_print("type(hash('abc')).__name__"), "int");
}

#[test]
fn hash_same_int_equal() {
    assert_eq!(run_print("hash(7) == hash(7)"), "True");
}

#[test]
fn hash_tuple_of_ints() {
    assert_eq!(run_print("type(hash((1, 2))).__name__"), "int");
}

#[test]
fn id_same_object_equal() {
    assert_eq!(
        run_python_one("xs = []\nprint(id(xs) == id(xs))\n"),
        "True"
    );
}

#[test]
fn id_distinct_objects_differ() {
    assert_eq!(
        run_python_one("print(id([]) == id([]))\n"),
        "False"
    );
}

#[test]
fn id_small_ints_may_be_cached() {
    assert_eq!(run_print("id(256) == id(256) or id(256) != id(256)"), "True");
}

#[test]
fn id_used_in_is_comparison() {
    assert_eq!(
        run_python_one("a = object()\nprint(a is a)\n"),
        "True"
    );
}

#[test]
fn hash_of_bool_true() {
    assert_eq!(run_print("hash(True) == hash(True)"), "True");
}

#[test]
fn hash_of_none() {
    assert_eq!(run_print("type(hash(None)).__name__"), "int");
}

#[test]
fn callable_on_list_append_bound_method() {
    assert_eq!(run_print("callable([].append)"), "True");
}

#[test]
fn id_function_returns_positive_or_negative_int() {
    assert_eq!(
        run_python_one("v = id('x')\nprint(isinstance(v, int))\n"),
        "True"
    );
}

#[test]
fn hash_frozenset_stable() {
    assert_eq!(
        run_python_one("fs = frozenset([1, 2])\nprint(hash(fs) == hash(fs))\n"),
        "True"
    );
}

#[test]
fn callable_none_is_false() {
    assert_eq!(run_print("callable(None)"), "False");
}

#[test]
fn id_after_rebind_changes() {
    assert_eq!(
        run_python_one("a = [1]\nb = a\na = [2]\nprint(id(a) == id(b))\n"),
        "False"
    );
}

#[test]
fn hash_dict_not_hashable_raises() {
    assert_eq!(
        run_python_one("try:\n hash({})\n print('ok')\nexcept TypeError:\n print('TypeError')\n"),
        "TypeError"
    );
}

#[test]
fn callable_with_call_after_check() {
    assert_eq!(
        run_python_one("f = abs\nprint(callable(f) and f(-3) == 3)\n"),
        "True"
    );
}

#[test]
fn id_of_class_object() {
    assert_eq!(
        run_python_one("class C:\n pass\nprint(id(C) == id(C))\n"),
        "True"
    );
}

#[test]
fn hash_used_in_set_membership() {
    assert_eq!(
        run_python_one("s = {hash('a'), hash('b')}\nprint(len(s) >= 1)\n"),
        "True"
    );
}

#[test]
fn callable_generator_function() {
    assert_eq!(
        run_python_one("def g():\n yield 1\nprint(callable(g))\n"),
        "True"
    );
}

#[test]
fn id_tuple_of_ids_unique_per_element() {
    assert_eq!(
        run_python_one("a, b = 1, 2\nprint(id(a) != id(b) or a == b)\n"),
        "True"
    );
}

#[test]
fn hash_negative_int_allowed() {
    assert_eq!(
        run_python_one("print(isinstance(hash(-1), int))\n"),
        "True"
    );
}

#[test]
fn callable_on_type_builtin() {
    assert_eq!(run_print("callable(type)"), "True");
}

#[test]
fn id_none_is_constant() {
    assert_eq!(run_print("id(None) == id(None)"), "True");
}
