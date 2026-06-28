use crate::helpers::run_in_main;

#[test]
fn subclass_overrides_parent_instance_method() {
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
fn override_dispatches_to_most_specific_implementation() {
    let types = r#"
        static class Shape { String label() { return "shape"; } }
        static class Circle extends Shape { String label() { return "circle"; } }
        static class Dot extends Circle { String label() { return "dot"; } }
    "#;
    let out = run_in_main("Shape s = new Dot(); System.out.println(s.label());", types);
    assert_eq!(out, vec!["dot"]);
}

#[test]
fn super_invokes_parent_method_body() {
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
fn super_constructor_initializes_parent_fields() {
    let types = r#"
        static class A { int x; A(int v) { x = v; } }
        static class B extends A { B(int v) { super(v); } }
    "#;
    let out = run_in_main("B b = new B(42); System.out.println(b.x);", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn child_inherits_parent_instance_field() {
    let types = r#"
        static class Parent { int value = 7; }
        static class Child extends Parent {}
    "#;
    let out = run_in_main("Child c = new Child(); System.out.println(c.value);", types);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn child_reads_protected_field_from_parent() {
    let types = r#"
        static class Base { protected int secret = 11; }
        static class Child extends Base {
            int readSecret() { return secret; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.readSecret());",
        types,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn child_calls_protected_parent_method() {
    let types = r#"
        static class Base { protected int reveal() { return 5; } }
        static class Child extends Base {
            int viaProtected() { return reveal(); }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.viaProtected());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn subclass_assigns_inherited_protected_field() {
    let types = r#"
        static class Base { protected int tally = 0; }
        static class Child extends Base {
            void bump() { tally = tally + 3; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); c.bump(); System.out.println(c.tally);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn three_level_inheritance_preserves_grandparent_field() {
    let types = r#"
        static class Grand { int depth = 1; }
        static class Parent extends Grand { int depth = 2; }
        static class Child extends Parent {}
    "#;
    let out = run_in_main("Child c = new Child(); System.out.println(c.depth);", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn parent_method_used_when_child_does_not_override() {
    let types = r#"
        static class Engine { int horsepower() { return 100; } }
        static class TurboEngine extends Engine {}
    "#;
    let out = run_in_main(
        "TurboEngine t = new TurboEngine(); System.out.println(t.horsepower());",
        types,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn child_overrides_one_method_and_inherits_another() {
    let types = r#"
        static class Account {
            String kind() { return "account"; }
            int balance() { return 0; }
        }
        static class Savings extends Account {
            String kind() { return "savings"; }
        }
    "#;
    let out = run_in_main(
        "Savings s = new Savings(); System.out.println(s.kind()); System.out.println(s.balance());",
        types,
    );
    assert_eq!(out, vec!["savings", "0"]);
}

#[test]
fn upcast_subclass_reference_to_parent_type() {
    let types = r#"
        static class Vehicle { String move() { return "go"; } }
        static class Car extends Vehicle { String move() { return "drive"; } }
    "#;
    let out = run_in_main(
        "Car car = new Car(); Vehicle v = car; System.out.println(v.move());",
        types,
    );
    assert_eq!(out, vec!["drive"]);
}

#[test]
fn upcast_preserves_dynamic_type_for_instanceof() {
    let types = r#"
        static class Fruit {}
        static class Apple extends Fruit {}
    "#;
    let out = run_in_main(
        "Fruit f = new Apple(); System.out.println(f instanceof Apple);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_true_for_exact_runtime_class() {
    let types = r#"
        static class Node {}
        static class Leaf extends Node {}
    "#;
    let out = run_in_main(
        "Leaf leaf = new Leaf(); System.out.println(leaf instanceof Leaf);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_true_for_supertype_of_instance() {
    let types = r#"
        static class Animal {}
        static class Cat extends Animal {}
    "#;
    let out = run_in_main(
        "Animal a = new Cat(); System.out.println(a instanceof Animal);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_false_for_unrelated_type() {
    let types = r#"
        static class Dog {}
        static class Tree {}
        static class Puppy extends Dog {}
    "#;
    let out = run_in_main(
        "Puppy p = new Puppy(); System.out.println(p instanceof Tree);",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instanceof_false_for_sibling_subclass() {
    let types = r#"
        static class Pet {}
        static class Cat extends Pet {}
        static class Dog extends Pet {}
    "#;
    let out = run_in_main(
        "Pet p = new Cat(); System.out.println(p instanceof Dog);",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn downcast_after_instanceof_reads_subclass_field() {
    let types = r#"
        static class Box { int size = 1; }
        static class BigBox extends Box { int size = 9; }
    "#;
    let out = run_in_main(
        "Box b = new BigBox(); if (b instanceof BigBox bb) { System.out.println(bb.size); }",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn super_call_from_overridden_void_method_runs_parent_logic() {
    let types = r#"
        static class Logger { void log(String msg) { System.out.println("parent:" + msg); } }
        static class AuditLogger extends Logger {
            void log(String msg) { super.log(msg); System.out.println("child:" + msg); }
        }
    "#;
    let out = run_in_main("AuditLogger a = new AuditLogger(); a.log(\"x\");", types);
    assert_eq!(out, vec!["parent:x", "child:x"]);
}

#[test]
fn child_constructor_can_invoke_parameterized_super_constructor() {
    let types = r#"
        static class Pair { int a; int b; Pair(int a, int b) { this.a = a; this.b = b; } }
        static class Triple extends Pair {
            int c;
            Triple(int a, int b, int c) { super(a, b); this.c = c; }
        }
    "#;
    let out = run_in_main(
        "Triple t = new Triple(1, 2, 3); System.out.println(t.a); System.out.println(t.c);",
        types,
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn inherited_method_visible_on_upcasted_reference() {
    let types = r#"
        static class Counter { int next() { return 1; } }
        static class StepCounter extends Counter { int next() { return 2; } }
    "#;
    let out = run_in_main(
        "StepCounter s = new StepCounter(); Counter c = s; System.out.println(c.next());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn field_shadowing_child_field_hides_parent_name() {
    let types = r#"
        static class Base { int value = 1; }
        static class Child extends Base { int value = 2; }
    "#;
    let out = run_in_main("Child c = new Child(); System.out.println(c.value);", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn super_accesses_shadowed_parent_field() {
    let types = r#"
        static class Base { int value = 1; }
        static class Child extends Base {
            int value = 2;
            int parentValue() { return super.value; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.parentValue());",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn abstract_parent_requires_concrete_child_implementation() {
    let types = r#"
        static abstract class Expr { abstract int eval(); }
        static class Literal extends Expr {
            int value;
            Literal(int value) { this.value = value; }
            int eval() { return value; }
        }
    "#;
    let out = run_in_main(
        "Expr e = new Literal(8); System.out.println(e.eval());",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn polymorphic_call_through_abstract_base_type() {
    let types = r#"
        static abstract class Op { abstract int apply(int x); }
        static class DoubleOp extends Op { int apply(int x) { return x * 2; } }
        static class IncOp extends Op { int apply(int x) { return x + 1; } }
    "#;
    let out = run_in_main(
        "Op d = new DoubleOp(); Op i = new IncOp(); System.out.println(d.apply(3)); System.out.println(i.apply(3));",
        types,
    );
    assert_eq!(out, vec!["6", "4"]);
}

#[test]
fn child_extends_parent_with_protected_mutator() {
    let types = r#"
        static class Cache { protected int hits = 0; protected void hit() { hits++; } }
        static class CountingCache extends Cache {
            int report() { hit(); hit(); return hits; }
        }
    "#;
    let out = run_in_main(
        "CountingCache c = new CountingCache(); System.out.println(c.report());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn multilevel_override_calls_middle_super() {
    let types = r#"
        static class A { String id() { return "A"; } }
        static class B extends A { String id() { return super.id() + "B"; } }
        static class C extends B { String id() { return super.id() + "C"; } }
    "#;
    let out = run_in_main("C c = new C(); System.out.println(c.id());", types);
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn upcast_array_of_subclass_to_parent_array_element() {
    let types = r#"
        static class Item { String name() { return "item"; } }
        static class Book extends Item { String name() { return "book"; } }
    "#;
    let out = run_in_main(
        "Book b = new Book(); Item i = b; System.out.println(i.name());",
        types,
    );
    assert_eq!(out, vec!["book"]);
}

#[test]
fn instanceof_on_null_reference_is_false() {
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
fn child_adds_method_alongside_inherited_api() {
    let types = r#"
        static class Base { int baseValue() { return 1; } }
        static class Extended extends Base { int extraValue() { return 10; } }
    "#;
    let out = run_in_main(
        "Extended e = new Extended(); System.out.println(e.baseValue() + e.extraValue());",
        types,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn protected_method_overridden_in_child_still_callable_via_super() {
    let types = r#"
        static class Base { protected int step() { return 1; } }
        static class Child extends Base {
            protected int step() { return super.step() + 2; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.step());",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn extends_chain_constructor_sets_all_levels() {
    let types = r#"
        static class L1 { int a = 1; }
        static class L2 extends L1 { int b = 2; }
        static class L3 extends L2 { int sum() { return a + b; } }
    "#;
    let out = run_in_main("L3 l = new L3(); System.out.println(l.sum());", types);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn runtime_type_preserved_after_upcast_assignment() {
    let types = r#"
        static class Printer { String output() { return "p"; } }
        static class FancyPrinter extends Printer { String output() { return "fp"; } }
    "#;
    let out = run_in_main(
        "Printer p = new FancyPrinter(); System.out.println(p instanceof FancyPrinter); System.out.println(p.output());",
        types,
    );
    assert_eq!(out, vec!["true", "fp"]);
}

#[test]
fn sibling_instances_do_not_share_override_state() {
    let types = r#"
        static class Counter { int n = 0; void inc() { n++; } }
        static class NamedCounter extends Counter { String name; NamedCounter(String name) { this.name = name; } }
    "#;
    let out = run_in_main(
        "NamedCounter a = new NamedCounter(\"a\"); NamedCounter b = new NamedCounter(\"b\"); a.inc(); System.out.println(a.n); System.out.println(b.n);",
        types,
    );
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn override_returning_string_coexists_with_parent_int_field() {
    let types = r#"
        static class Record { int id = 5; String describe() { return "old"; } }
        static class NewRecord extends Record { String describe() { return "new-" + id; } }
    "#;
    let out = run_in_main(
        "Record r = new NewRecord(); System.out.println(r.describe());",
        types,
    );
    assert_eq!(out, vec!["new-5"]);
}
