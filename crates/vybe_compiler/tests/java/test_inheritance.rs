use crate::helpers::run_in_main;

#[test]
fn subclass_overrides_superclass_method() {
    let types = r#"
        static class Animal { String speak() { return "..."; } }
        static class Dog extends Animal { String speak() { return "woof"; } }
    "#;
    let out = run_in_main(
        "Animal a = new Dog(); System.out.println(a.speak());",
        types,
    );
    assert_eq!(out, vec!["woof"]);
}

#[test]
fn super_calls_parent_method_implementation() {
    let types = r#"
        static class Base { String tag() { return "base"; } }
        static class Derived extends Base {
            String tag() { return super.tag() + "-derived"; }
        }
    "#;
    let out = run_in_main(
        "Derived d = new Derived(); System.out.println(d.tag());",
        types,
    );
    assert_eq!(out, vec!["base-derived"]);
}

#[test]
fn subclass_inherits_parent_field() {
    let types = r#"
        static class Parent { int value = 7; }
        static class Child extends Parent {}
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.value);",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn subclass_constructor_chains_to_super() {
    let types = r#"
        static class A { int x; A(int v) { x = v; } }
        static class B extends A { B(int v) { super(v); } }
    "#;
    let out = run_in_main(
        "B b = new B(42); System.out.println(b.x);",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn instanceof_detects_subclass_instance() {
    let types = r#"
        static class Animal {}
        static class Cat extends Animal {}
    "#;
    let out = run_in_main(
        "Animal a = new Cat(); System.out.println(a instanceof Cat); System.out.println(a instanceof Animal);",
        types,
    );
    assert_eq!(out, vec!["true", "true"]);
}
