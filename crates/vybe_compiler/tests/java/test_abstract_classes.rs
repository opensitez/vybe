use crate::helpers::{compile_ok_check, run_in_main};

#[test]
fn concrete_subclass_implements_single_abstract_method() {
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
fn abstract_method_dispatched_polymorphically() {
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
fn abstract_class_direct_instantiation_does_not_compile() {
    let src = r#"
        public class Main {
            static abstract class Base { abstract int value(); }
            public static void main(String[] args) {
                Base b = new Base();
            }
        }
    "#;
    assert!(!compile_ok_check(src));
}

#[test]
fn abstract_class_with_concrete_methods_reused() {
    let types = r#"
        static abstract class Base {
            int twice(int n) { return n * 2; }
            abstract int core();
        }
        static class Child extends Base { int core() { return 3; } }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.twice(c.core()));",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn abstract_class_with_fields_initialized_in_subclass() {
    let types = r#"
        static abstract class Holder { int value; }
        static class Box extends Holder { Box(int v) { value = v; } }
    "#;
    let out = run_in_main(
        "Box b = new Box(15); System.out.println(b.value);",
        types,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn abstract_subclass_still_requires_concrete_grandchild() {
    let types = r#"
        static abstract class A { abstract int n(); }
        static abstract class B extends A {}
        static class C extends B { int n() { return 7; } }
    "#;
    let out = run_in_main(
        "A a = new C(); System.out.println(a.n());",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn template_method_pattern_via_abstract_hook() {
    let types = r#"
        static abstract class Template {
            int run() { return hook() + 1; }
            abstract int hook();
        }
        static class Impl extends Template { int hook() { return 4; } }
    "#;
    let out = run_in_main(
        "Template t = new Impl(); System.out.println(t.run());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn abstract_parent_constructor_called_from_child() {
    let types = r#"
        static abstract class Base {
            int seed;
            Base(int seed) { this.seed = seed; }
            abstract int next();
        }
        static class Child extends Base {
            Child(int seed) { super(seed); }
            int next() { return seed + 1; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(10); System.out.println(c.next());",
        types,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn abstract_method_overridden_in_grandchild() {
    let types = r#"
        static abstract class A { abstract String id(); }
        static abstract class B extends A {}
        static class C extends B { String id() { return "C"; } }
    "#;
    let out = run_in_main(
        "A a = new C(); System.out.println(a.id());",
        types,
    );
    assert_eq!(out, vec!["C"]);
}

#[test]
fn abstract_class_with_protected_abstract_method() {
    let types = r#"
        static abstract class Base { protected abstract int secret(); }
        static class Child extends Base {
            protected int secret() { return 12; }
            int reveal() { return secret(); }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.reveal());",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn abstract_class_static_method_not_overridden_by_instance() {
    let types = r#"
        static abstract class Base { static int code() { return 1; } abstract int inst(); }
        static class Child extends Base { int inst() { return 2; } }
    "#;
    let out = run_in_main(
        "System.out.println(Base.code()); Child c = new Child(); System.out.println(c.inst());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn multiple_concrete_implementations_of_same_abstract() {
    let types = r#"
        static abstract class Codec { abstract String encode(int n); }
        static class HexCodec extends Codec { String encode(int n) { return "h" + n; } }
        static class DecCodec extends Codec { String encode(int n) { return "d" + n; } }
    "#;
    let out = run_in_main(
        "Codec h = new HexCodec(); Codec d = new DecCodec(); System.out.println(h.encode(1)); System.out.println(d.encode(1));",
        types,
    );
    assert_eq!(out, vec!["h1", "d1"]);
}

#[test]
fn abstract_method_returns_string_from_subclass() {
    let types = r#"
        static abstract class Labeler { abstract String label(); }
        static class Simple extends Labeler { String label() { return "ok"; } }
    "#;
    let out = run_in_main(
        "Labeler l = new Simple(); System.out.println(l.label());",
        types,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn abstract_class_with_final_concrete_method() {
    let types = r#"
        static abstract class Base {
            final int fixed() { return 99; }
            abstract int var();
        }
        static class Child extends Base { int var() { return 1; } }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.fixed()); System.out.println(c.var());",
        types,
    );
    assert_eq!(out, vec!["99", "1"]);
}

#[test]
fn abstract_base_with_default_behavior_in_concrete() {
    let types = r#"
        static abstract class Parser { abstract boolean valid(String s); }
        static class NonEmpty extends Parser { boolean valid(String s) { return s.length() > 0; } }
    "#;
    let out = run_in_main(
        "Parser p = new NonEmpty(); System.out.println(p.valid(\"x\")); System.out.println(p.valid(\"\"));",
        types,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn abstract_class_field_accessible_in_subclass() {
    let types = r#"
        static abstract class Base { protected int tally = 0; }
        static class Child extends Base {
            void bump() { tally++; }
            int read() { return tally; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); c.bump(); c.bump(); System.out.println(c.read());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn abstract_parent_and_interface_combo() {
    let types = r#"
        static interface Named { String name(); }
        static abstract class Entity implements Named { abstract int id(); }
        static class User extends Entity {
            String name() { return "user"; }
            int id() { return 1; }
        }
    "#;
    let out = run_in_main(
        "Entity e = new User(); System.out.println(e.name()); System.out.println(e.id());",
        types,
    );
    assert_eq!(out, vec!["user", "1"]);
}

#[test]
fn abstract_method_takes_parameters() {
    let types = r#"
        static abstract class MathFn { abstract int eval(int a, int b); }
        static class Add extends MathFn { int eval(int a, int b) { return a + b; } }
    "#;
    let out = run_in_main(
        "MathFn fn = new Add(); System.out.println(fn.eval(4, 5));",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn abstract_method_uses_subclass_field() {
    let types = r#"
        static abstract class Base { abstract int compute(); }
        static class Scaled extends Base {
            int factor = 3;
            int compute() { return factor * 4; }
        }
    "#;
    let out = run_in_main(
        "Base b = new Scaled(); System.out.println(b.compute());",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn sibling_subclasses_independent_state() {
    let types = r#"
        static abstract class Counter { int n = 0; abstract void inc(); }
        static class A extends Counter { void inc() { n++; } }
        static class B extends Counter { void inc() { n = n + 2; } }
    "#;
    let out = run_in_main(
        "A a = new A(); B b = new B(); a.inc(); b.inc(); System.out.println(a.n); System.out.println(b.n);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn abstract_class_with_parameterized_constructor() {
    let types = r#"
        static abstract class Node { int depth; Node(int depth) { this.depth = depth; } abstract int size(); }
        static class Leaf extends Node {
            Leaf(int depth) { super(depth); }
            int size() { return 1; }
        }
    "#;
    let out = run_in_main(
        "Node n = new Leaf(2); System.out.println(n.depth); System.out.println(n.size());",
        types,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn abstract_method_chain_super_concrete() {
    let types = r#"
        static abstract class Base { int base() { return 1; } abstract int total(); }
        static class Child extends Base { int total() { return base() + 2; } }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.total());",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn abstract_shape_hierarchy_area() {
    let types = r#"
        static abstract class Shape { abstract int area(); }
        static class Square extends Shape {
            int side;
            Square(int side) { this.side = side; }
            int area() { return side * side; }
        }
    "#;
    let out = run_in_main(
        "Shape s = new Square(4); System.out.println(s.area());",
        types,
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn abstract_animal_speak_polymorphism() {
    let types = r#"
        static abstract class Animal { abstract String speak(); }
        static class Dog extends Animal { String speak() { return "woof"; } }
        static class Cat extends Animal { String speak() { return "meow"; } }
    "#;
    let out = run_in_main(
        "Animal d = new Dog(); Animal c = new Cat(); System.out.println(d.speak()); System.out.println(c.speak());",
        types,
    );
    assert_eq!(out, vec!["woof", "meow"]);
}

#[test]
fn abstract_expression_eval_add() {
    let types = r#"
        static abstract class Expr { abstract int eval(); }
        static class Add extends Expr {
            Expr left;
            Expr right;
            Add(Expr left, Expr right) { this.left = left; this.right = right; }
            int eval() { return left.eval() + right.eval(); }
        }
        static class Num extends Expr {
            int value;
            Num(int value) { this.value = value; }
            int eval() { return value; }
        }
    "#;
    let out = run_in_main(
        "Expr e = new Add(new Num(2), new Num(3)); System.out.println(e.eval());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn abstract_expression_eval_literal() {
    let types = r#"
        static abstract class Expr { abstract int eval(); }
        static class Lit extends Expr {
            int value;
            Lit(int value) { this.value = value; }
            int eval() { return value; }
        }
    "#;
    let out = run_in_main(
        "Expr e = new Lit(42); System.out.println(e.eval());",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn partial_abstract_middle_class() {
    let types = r#"
        static abstract class Root { abstract int a(); }
        static abstract class Mid extends Root { abstract int b(); }
        static class Leaf extends Mid { int a() { return 1; } int b() { return 2; } }
    "#;
    let out = run_in_main(
        "Root r = new Leaf(); System.out.println(r.a());",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn concrete_extends_partial_abstract() {
    let types = r#"
        static abstract class Mid { abstract int both(); }
        static class Leaf extends Mid { int both() { return 5; } }
    "#;
    let out = run_in_main(
        "Mid m = new Leaf(); System.out.println(m.both());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn abstract_class_with_private_concrete_helper() {
    let types = r#"
        static abstract class Base {
            private int helper() { return 3; }
            abstract int expose();
        }
        static class Child extends Base {
            int expose() { return helper(); }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.expose());",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn abstract_method_return_type_primitive() {
    let types = r#"
        static abstract class Reader { abstract int read(); }
        static class Const extends Reader { int read() { return 7; } }
    "#;
    let out = run_in_main(
        "Reader r = new Const(); System.out.println(r.read());",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn abstract_void_method_side_effect() {
    let types = r#"
        static abstract class Logger { abstract void log(String msg); }
        static class PrintLogger extends Logger { void log(String msg) { System.out.println(msg); } }
    "#;
    let out = run_in_main(
        "Logger l = new PrintLogger(); l.log(\"trace\");",
        types,
    );
    assert_eq!(out, vec!["trace"]);
}

#[test]
fn abstract_class_instanceof_check() {
    let types = r#"
        static abstract class Base {}
        static class Child extends Base {}
    "#;
    let out = run_in_main(
        "Base b = new Child(); System.out.println(b instanceof Child);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn abstract_reference_holds_subclass() {
    let types = r#"
        static abstract class Vehicle { abstract String type(); }
        static class Car extends Vehicle { String type() { return "car"; } }
    "#;
    let out = run_in_main(
        "Vehicle v = new Car(); System.out.println(v.type());",
        types,
    );
    assert_eq!(out, vec!["car"]);
}

#[test]
fn abstract_method_called_from_concrete_in_same_class() {
    let types = r#"
        static abstract class Base {
            abstract int core();
            int wrapped() { return core() + 10; }
        }
        static class Child extends Base { int core() { return 2; } }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.wrapped());",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn abstract_game_update_loop_pattern() {
    let types = r#"
        static abstract class Game {
            int ticks = 0;
            abstract void step();
            void update() { step(); ticks++; }
        }
        static class SimpleGame extends Game { void step() {} }
    "#;
    let out = run_in_main(
        "SimpleGame g = new SimpleGame(); g.update(); g.update(); System.out.println(g.ticks);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn abstract_parser_parse_implementation() {
    let types = r#"
        static abstract class Parser { abstract int parse(String s); }
        static class IntParser extends Parser { int parse(String s) { return Integer.parseInt(s); } }
    "#;
    let out = run_in_main(
        "Parser p = new IntParser(); System.out.println(p.parse(\"21\"));",
        types,
    );
    assert_eq!(out, vec!["21"]);
}

#[test]
fn abstract_builder_build_method() {
    let types = r#"
        static abstract class Builder { abstract String build(); }
        static class GreetingBuilder extends Builder { String build() { return "hello"; } }
    "#;
    let out = run_in_main(
        "Builder b = new GreetingBuilder(); System.out.println(b.build());",
        types,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn abstract_class_with_static_factory_returns_concrete() {
    let types = r#"
        static abstract class Service { abstract int run(); }
        static class FastService extends Service { int run() { return 1; } }
        static class Services {
            static Service fast() { return new FastService(); }
        }
    "#;
    let out = run_in_main(
        "Service s = Services.fast(); System.out.println(s.run());",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn abstract_list_add_size_template() {
    let types = r#"
        static abstract class Bag { abstract int size(); abstract void add(int n); }
        static class IntBag extends Bag {
            int count = 0;
            int size() { return count; }
            void add(int n) { count = count + n; }
        }
    "#;
    let out = run_in_main(
        "Bag bag = new IntBag(); bag.add(2); bag.add(3); System.out.println(bag.size());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn abstract_node_depth_calculation() {
    let types = r#"
        static abstract class Node { abstract int depth(); }
        static class Leaf extends Node { int depth() { return 0; } }
        static class Branch extends Node {
            Node child;
            Branch(Node child) { this.child = child; }
            int depth() { return child.depth() + 1; }
        }
    "#;
    let out = run_in_main(
        "Node tree = new Branch(new Leaf()); System.out.println(tree.depth());",
        types,
    );
    assert_eq!(out, vec!["1"]);
}
