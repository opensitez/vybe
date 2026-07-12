use crate::helpers::{run_in_main, run_main};

#[test]
fn class_is_instance_true_for_same_class_object() {
    let out = run_main(r#"String s = "hello"; System.out.println(String.class.isInstance(s));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_false_for_unrelated_type() {
    let out = run_main(r#"Integer n = 5; System.out.println(String.class.isInstance(n));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_is_instance_true_for_subclass_via_superclass_ref() {
    let types = r#"
        static class Animal {}
        static class Dog extends Animal {}
    "#;
    let out = run_in_main(
        "Dog d = new Dog(); System.out.println(Animal.class.isInstance(d));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_false_superclass_for_subclass_check() {
    let types = r#"
        static class Animal {}
        static class Dog extends Animal {}
    "#;
    let out = run_in_main(
        "Animal a = new Animal(); System.out.println(Dog.class.isInstance(a));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_is_instance_true_for_interface_implementation() {
    let types = r#"
        interface Runnable2 { void run(); }
        static class Task implements Runnable2 { public void run() {} }
    "#;
    let out = run_in_main(
        "Task t = new Task(); System.out.println(Runnable2.class.isInstance(t));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_false_for_null_reference() {
    let out = run_main(r#"String s = null; System.out.println(String.class.isInstance(s));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_is_instance_true_for_boxed_integer() {
    let out = run_main(
        r#"Object o = Integer.valueOf(10); System.out.println(Integer.class.isInstance(o));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_true_integer_for_number_supertype() {
    let out = run_main(
        r#"Object o = Integer.valueOf(10); System.out.println(Number.class.isInstance(o));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_array_type_for_int_array() {
    let out =
        run_main(r#"int[] arr = new int[3]; System.out.println(int[].class.isInstance(arr));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_array_false_for_wrong_type() {
    let out =
        run_main(r#"int[] arr = new int[3]; System.out.println(String[].class.isInstance(arr));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_is_instance_object_class_accepts_any_nonnull() {
    let out =
        run_main(r#"Object o = new Object(); System.out.println(Object.class.isInstance(o));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_list_for_arraylist() {
    let out = run_main(
        r#"java.util.ArrayList<String> list = new java.util.ArrayList<String>(); System.out.println(java.util.List.class.isInstance(list));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_arraylist_false_for_linked_list() {
    let out = run_main(
        r#"java.util.LinkedList<String> list = new java.util.LinkedList<String>(); System.out.println(java.util.ArrayList.class.isInstance(list));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_is_instance_string_builder_instance() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); System.out.println(StringBuilder.class.isInstance(sb));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_string_builder_false_for_string() {
    let out =
        run_main(r#"String s = "text"; System.out.println(StringBuilder.class.isInstance(s));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_is_instance_exception_for_runtime_exception() {
    let out = run_main(
        r#"RuntimeException e = new RuntimeException(); System.out.println(Exception.class.isInstance(e));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_throwable_for_error() {
    let out = run_main(
        r#"OutOfMemoryError e = new OutOfMemoryError(); System.out.println(Throwable.class.isInstance(e));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_comparable_for_string() {
    let out = run_main(r#"String s = "sort"; System.out.println(Comparable.class.isInstance(s));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_enum_constant() {
    let types = r#"enum Color { RED, GREEN }"#;
    let out = run_in_main(
        "System.out.println(Color.class.isInstance(Color.RED));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_enum_false_for_string() {
    let types = r#"enum Color { RED }"#;
    let out = run_in_main(
        r#"System.out.println(Color.class.isInstance("RED"));"#,
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_is_instance_deep_hierarchy_grandchild() {
    let types = r#"
        static class A {}
        static class B extends A {}
        static class C extends B {}
    "#;
    let out = run_in_main(
        "C c = new C(); System.out.println(A.class.isInstance(c));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_deep_hierarchy_not_reverse() {
    let types = r#"
        static class A {}
        static class B extends A {}
        static class C extends B {}
    "#;
    let out = run_in_main(
        "A a = new A(); System.out.println(C.class.isInstance(a));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_is_instance_autoboxed_boolean() {
    let out =
        run_main(r#"Object o = Boolean.TRUE; System.out.println(Boolean.class.isInstance(o));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_autoboxed_character() {
    let out = run_main(
        r#"Object o = Character.valueOf('Z'); System.out.println(Character.class.isInstance(o));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_map_for_hashmap() {
    let out = run_main(
        r#"java.util.HashMap<String, Integer> m = new java.util.HashMap<String, Integer>(); System.out.println(java.util.Map.class.isInstance(m));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_set_for_hashset() {
    let out = run_main(
        r#"java.util.HashSet<Integer> s = new java.util.HashSet<Integer>(); System.out.println(java.util.Set.class.isInstance(s));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_collection_for_vector() {
    let out = run_main(
        r#"java.util.Vector<Integer> v = new java.util.Vector<Integer>(); System.out.println(java.util.Collection.class.isInstance(v));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_serializable_for_string() {
    let out = run_main(
        r#"String s = "ser"; System.out.println(java.io.Serializable.class.isInstance(s));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_cloneable_for_arraylist() {
    let out = run_main(
        r#"java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); System.out.println(Cloneable.class.isInstance(list));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_class_object_itself() {
    let out = run_main(r#"System.out.println(Class.class.isInstance(String.class));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_used_in_conditional_branch() {
    let out = run_main(
        r#"Object o = "data"; if (String.class.isInstance(o)) { System.out.println("str"); } else { System.out.println("other"); }"#,
    );
    assert_eq!(out, vec!["str"]);
}

#[test]
fn class_is_instance_used_in_conditional_else_branch() {
    let out = run_main(
        r#"Object o = 42; if (String.class.isInstance(o)) { System.out.println("str"); } else { System.out.println("other"); }"#,
    );
    assert_eq!(out, vec!["other"]);
}

#[test]
fn class_is_instance_double_wrapper() {
    let out = run_main(
        r#"Object o = Double.valueOf(3.14); System.out.println(Double.class.isInstance(o));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_long_wrapper() {
    let out =
        run_main(r#"Object o = Long.valueOf(100L); System.out.println(Long.class.isInstance(o));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_is_instance_byte_array_primitive() {
    let out = run_main(
        r#"byte[] data = new byte[4]; System.out.println(byte[].class.isInstance(data));"#,
    );
    assert_eq!(out, vec!["true"]);
}
