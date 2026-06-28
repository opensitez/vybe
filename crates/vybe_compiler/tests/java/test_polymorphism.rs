use crate::helpers::{run_in_main, run_main};

#[test]
fn upcast_subclass_to_parent_reference() {
    let types = r#"
        static class Animal { String speak() { return "..."; } }
        static class Dog extends Animal { String speak() { return "woof"; } }
    "#;
    let out = run_in_main(
        "Dog d = new Dog(); Animal a = d; System.out.println(a.speak());",
        types,
    );
    assert_eq!(out, vec!["woof"]);
}

#[test]
fn virtual_dispatch_calls_subclass_method() {
    let types = r#"
        static class Shape { String label() { return "shape"; } }
        static class Circle extends Shape { String label() { return "circle"; } }
    "#;
    let out = run_in_main(
        "Shape s = new Circle(); System.out.println(s.label());",
        types,
    );
    assert_eq!(out, vec!["circle"]);
}

#[test]
fn virtual_dispatch_through_grandparent_reference() {
    let types = r#"
        static class A { String id() { return "A"; } }
        static class B extends A { String id() { return "B"; } }
        static class C extends B { String id() { return "C"; } }
    "#;
    let out = run_in_main("A ref = new C(); System.out.println(ref.id());", types);
    assert_eq!(out, vec!["C"]);
}

#[test]
fn downcast_with_instanceof_pattern() {
    let types = r#"
        static class Base {}
        static class Child extends Base { int value = 7; }
    "#;
    let out = run_in_main(
        "Base b = new Child(); if (b instanceof Child c) { System.out.println(c.value); }",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn downcast_accesses_subclass_only_method() {
    let types = r#"
        static class Vehicle { String type() { return "vehicle"; } }
        static class Car extends Vehicle { String plate() { return "ABC"; } }
    "#;
    let out = run_in_main(
        "Vehicle v = new Car(); if (v instanceof Car c) { System.out.println(c.plate()); }",
        types,
    );
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn instanceof_false_for_unrelated_type() {
    let types = r#"
        static class Cat {}
        static class Dog {}
    "#;
    let out = run_in_main(
        "Cat c = new Cat(); System.out.println(c instanceof Dog);",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instanceof_true_for_exact_type() {
    let types = r#"
        static class Box {}
    "#;
    let out = run_in_main(
        "Box b = new Box(); System.out.println(b instanceof Box);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_true_for_subclass_of_declared_type() {
    let types = r#"
        static class Parent {}
        static class Child extends Parent {}
    "#;
    let out = run_in_main(
        "Parent p = new Child(); System.out.println(p instanceof Child);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn field_hiding_parent_field_unchanged_on_child() {
    let types = r#"
        static class Parent { int value = 1; }
        static class Child extends Parent { int value = 2; }
    "#;
    let out = run_in_main(
        "Child c = new Child(); Parent p = c; System.out.println(p.value); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn method_overriding_changes_runtime_dispatch() {
    let types = r#"
        static class Parent { String tag() { return "parent"; } }
        static class Child extends Parent { String tag() { return "child"; } }
    "#;
    let out = run_in_main(
        "Parent p = new Child(); System.out.println(p.tag());",
        types,
    );
    assert_eq!(out, vec!["child"]);
}

#[test]
fn upcast_preserves_runtime_type_for_instanceof() {
    let types = r#"
        static class Printer { String output() { return "p"; } }
        static class FancyPrinter extends Printer { String output() { return "fp"; } }
    "#;
    let out = run_in_main(
        "Printer p = new FancyPrinter(); System.out.println(p instanceof FancyPrinter);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn polymorphic_array_of_parent_holds_subclasses() {
    let types = r#"
        static class Item { String name() { return "item"; } }
        static class Book extends Item { String name() { return "book"; } }
    "#;
    let out = run_in_main(
        "Item[] arr = new Item[] { new Book() }; System.out.println(arr[0].name());",
        types,
    );
    assert_eq!(out, vec!["book"]);
}

#[test]
fn virtual_dispatch_in_loop_over_mixed_types() {
    let types = r#"
        static class Node { int val() { return 0; } }
        static class A extends Node { int val() { return 1; } }
        static class B extends Node { int val() { return 2; } }
    "#;
    let out = run_in_main(
        "Node[] nodes = new Node[] { new A(), new B() }; int sum = 0; for (int i = 0; i < nodes.length; i++) { sum = sum + nodes[i].val(); } System.out.println(sum);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn field_access_via_parent_ref_uses_declared_type_field() {
    let types = r#"
        static class Base { int n = 10; }
        static class Derived extends Base { int n = 20; }
    "#;
    let out = run_in_main(
        "Derived d = new Derived(); Base b = d; System.out.println(b.n);",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn method_call_via_parent_ref_uses_runtime_type_method() {
    let types = r#"
        static class Base { int read() { return 1; } }
        static class Derived extends Base { int read() { return 2; } }
    "#;
    let out = run_in_main(
        "Derived d = new Derived(); Base b = d; System.out.println(b.read());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn super_method_not_called_when_overridden() {
    let types = r#"
        static class Base { String id() { return "base"; } }
        static class Child extends Base { String id() { return "child"; } }
    "#;
    let out = run_in_main("Child c = new Child(); System.out.println(c.id());", types);
    assert_eq!(out, vec!["child"]);
}

#[test]
fn super_method_called_explicitly_from_override() {
    let types = r#"
        static class Base { String id() { return "base"; } }
        static class Child extends Base { String id() { return super.id() + "-child"; } }
    "#;
    let out = run_in_main("Child c = new Child(); System.out.println(c.id());", types);
    assert_eq!(out, vec!["base-child"]);
}

#[test]
fn downcast_after_instanceof_guard() {
    let types = r#"
        static class Shape {}
        static class Circle extends Shape { int radius = 5; }
    "#;
    let out = run_in_main(
        "Shape s = new Circle(); if (s instanceof Circle) { Circle c = (Circle) s; System.out.println(c.radius); }",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn cross_hierarchy_instanceof_false() {
    let types = r#"
        static class A {}
        static class B {}
    "#;
    let out = run_in_main("A a = new A(); System.out.println(a instanceof B);", types);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn interface_reference_virtual_dispatch() {
    let types = r#"
        static interface Greeter { String greet(); }
        static class En implements Greeter { String greet() { return "hi"; } }
    "#;
    let out = run_in_main(
        "Greeter g = new En(); System.out.println(g.greet());",
        types,
    );
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn abstract_reference_virtual_dispatch() {
    let types = r#"
        static abstract class Op { abstract int apply(int x); }
        static class DoubleOp extends Op { int apply(int x) { return x * 2; } }
    "#;
    let out = run_in_main(
        "Op op = new DoubleOp(); System.out.println(op.apply(4));",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn double_dispatch_pattern_two_levels() {
    let types = r#"
        static class A { String id() { return "A"; } }
        static class B extends A { String id() { return super.id() + "B"; } }
        static class C extends B { String id() { return super.id() + "C"; } }
    "#;
    let out = run_in_main("A ref = new C(); System.out.println(ref.id());", types);
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn upcast_then_downcast_roundtrip() {
    let types = r#"
        static class Base {}
        static class Child extends Base { int code() { return 9; } }
    "#;
    let out = run_in_main(
        "Child c1 = new Child(); Base b = c1; Child c2 = (Child) b; System.out.println(c2.code());",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn instanceof_on_null_is_false() {
    let types = r#"
        static class Thing {}
    "#;
    let out = run_in_main(
        "Thing t = null; System.out.println(t instanceof Thing);",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn field_hiding_with_super_dot_access() {
    let types = r#"
        static class Base { int value = 1; }
        static class Child extends Base {
            int value = 2;
            int parentValue() { return super.value; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.parentValue()); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn static_method_not_polymorphic_via_reference() {
    let types = r#"
        static class Base { static int code() { return 1; } }
        static class Child extends Base { static int code() { return 2; } }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(Base.code()); System.out.println(Child.code());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn sibling_types_instanceof_distinguishes() {
    let types = r#"
        static class Parent {}
        static class Left extends Parent {}
        static class Right extends Parent {}
    "#;
    let out = run_in_main(
        "Parent p = new Left(); System.out.println(p instanceof Left); System.out.println(p instanceof Right);",
        types,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn chained_upcasts_same_dispatch() {
    let types = r#"
        static class L1 { String tag() { return "1"; } }
        static class L2 extends L1 { String tag() { return "2"; } }
        static class L3 extends L2 { String tag() { return "3"; } }
    "#;
    let out = run_in_main(
        "L3 obj = new L3(); L1 ref = obj; System.out.println(ref.tag());",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn polymorphic_to_string_override() {
    let types = r#"
        static class Base { String toString() { return "base"; } }
        static class Child extends Base { String toString() { return "child"; } }
    "#;
    let out = run_in_main(
        "Base b = new Child(); System.out.println(b.toString());",
        types,
    );
    assert_eq!(out, vec!["child"]);
}

#[test]
fn override_with_wider_behavior() {
    let types = r#"
        static class Base { int step() { return 1; } }
        static class Child extends Base { int step() { return super.step() + 2; } }
    "#;
    let out = run_in_main("Base b = new Child(); System.out.println(b.step());", types);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn parent_field_shadow_child_field_distinct() {
    let types = r#"
        static class Grand { int depth = 1; }
        static class Parent extends Grand { int depth = 2; }
        static class Child extends Parent {}
    "#;
    let out = run_in_main("Child c = new Child(); System.out.println(c.depth);", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn virtual_method_called_from_parent_method() {
    let types = r#"
        static class Base {
            String wrap() { return "[" + core() + "]"; }
            String core() { return "base"; }
        }
        static class Child extends Base { String core() { return "child"; } }
    "#;
    let out = run_in_main("Base b = new Child(); System.out.println(b.wrap());", types);
    assert_eq!(out, vec!["[child]"]);
}

#[test]
fn pattern_instanceof_string_in_guard() {
    let out = run_main(
        "Object o = \"java\"; if (o instanceof String s) { System.out.println(s.toUpperCase()); }",
    );
    assert_eq!(out, vec!["JAVA"]);
}

#[test]
fn array_instanceof_check() {
    let out = run_main("Object o = new int[3]; System.out.println(o instanceof int[]);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn polymorphic_equals_uses_override() {
    let types = r#"
        static class Base {
            int id;
            Base(int id) { this.id = id; }
            boolean equals(Object o) {
                if (o instanceof Base) { return ((Base) o).id == id; }
                return false;
            }
        }
        static class Child extends Base { Child(int id) { super(id); } }
    "#;
    let out = run_in_main(
        "Base a = new Child(1); Base b = new Child(1); System.out.println(a.equals(b));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn interface_instanceof_on_implementor() {
    let types = r#"
        static interface Closeable { void close(); }
        static class FileHandle implements Closeable { void close() {} }
    "#;
    let out = run_in_main(
        "Closeable c = new FileHandle(); System.out.println(c instanceof FileHandle);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn multilevel_override_deepest_wins() {
    let types = r#"
        static class A { int n() { return 1; } }
        static class B extends A { int n() { return 2; } }
        static class C extends B { int n() { return 3; } }
    "#;
    let out = run_in_main("A ref = new C(); System.out.println(ref.n());", types);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn field_hiding_vs_overriding_combined_output() {
    let types = r#"
        static class Base {
            int value = 1;
            int read() { return value; }
        }
        static class Child extends Base {
            int value = 2;
            int read() { return value + 10; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); Base b = c; System.out.println(b.value); System.out.println(c.value); System.out.println(b.read()); System.out.println(c.read());",
        types,
    );
    assert_eq!(out, vec!["1", "2", "12", "12"]);
}

#[test]
fn upcast_assignment_then_method_overload_resolution() {
    let types = r#"
        static class Parent { String kind() { return "parent"; } }
        static class Child extends Parent { String kind() { return "child"; } }
    "#;
    let out = run_in_main(
        "Parent p = new Child(); Child c = new Child(); System.out.println(p.kind()); System.out.println(c.kind());",
        types,
    );
    assert_eq!(out, vec!["child", "child"]);
}

#[test]
fn instanceof_parent_true_for_child_instance() {
    let types = r#"
        static class Parent {}
        static class Child extends Parent {}
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c instanceof Parent);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}
