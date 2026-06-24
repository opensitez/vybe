use crate::helpers::{run_in_main, run_main};

#[test]
fn final_local_int_retains_assigned_value() {
    let out = run_main("final int x = 10; System.out.println(x);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn final_local_string_reference_is_immutable_binding() {
    let out = run_main(r#"final String s = "locked"; System.out.println(s);"#);
    assert_eq!(out, vec!["locked"]);
}

#[test]
fn final_local_used_in_arithmetic_expression() {
    let out = run_main("final int base = 7; System.out.println(base * 2);");
    assert_eq!(out, vec!["14"]);
}

#[test]
fn final_local_passed_to_method_call() {
    let out = run_main("final int n = 5; System.out.println(Integer.toString(n));");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn final_field_initialized_at_declaration() {
    let types = r#"
        static class Box {
            final int size = 3;
        }
    "#;
    let out = run_in_main("Box b = new Box(); System.out.println(b.size);", types);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn final_field_string_literal_initializer() {
    let types = r#"
        static class Tag {
            final String label = "core";
        }
    "#;
    let out = run_in_main("Tag t = new Tag(); System.out.println(t.label);", types);
    assert_eq!(out, vec!["core"]);
}

#[test]
fn final_blank_field_set_in_constructor() {
    let types = r#"
        static class Counter {
            final int value;
            Counter(int v) { value = v; }
        }
    "#;
    let out = run_in_main("Counter c = new Counter(9); System.out.println(c.value);", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn final_static_field_shared_across_instances() {
    let types = r#"
        static class Config {
            static final int MAX = 100;
        }
    "#;
    let out = run_in_main("System.out.println(Config.MAX);", types);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn final_static_string_constant() {
    let types = r#"
        static class Names {
            static final String APP = "vybe";
        }
    "#;
    let out = run_in_main("System.out.println(Names.APP);", types);
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn final_method_in_parent_invoked_from_child_reference() {
    let types = r#"
        static class Base {
            final int fixed() { return 42; }
        }
        static class Child extends Base {}
    "#;
    let out = run_in_main(
        "Base b = new Child(); System.out.println(b.fixed());",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn final_method_called_on_child_instance_directly() {
    let types = r#"
        static class Base {
            final String tag() { return "final"; }
        }
        static class Child extends Base {}
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.tag());",
        types,
    );
    assert_eq!(out, vec!["final"]);
}

#[test]
fn final_method_not_overridden_by_child_adds_own_method() {
    let types = r#"
        static class Base {
            final int baseId() { return 1; }
        }
        static class Child extends Base {
            int childId() { return 2; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.baseId()); System.out.println(c.childId());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn final_class_instance_created_and_used() {
    let types = r#"
        final class Seal {}
    "#;
    let out = run_in_main(
        "Seal s = new Seal(); System.out.println(s != null);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn final_class_with_final_field_and_method() {
    let types = r#"
        final class Token {
            final int code = 7;
            final int code() { return code; }
        }
    "#;
    let out = run_in_main(
        "Token t = new Token(); System.out.println(t.code());",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn final_parameter_treated_as_local_binding() {
    let types = r#"
        static int doubleIt(final int n) { return n * 2; }
    "#;
    let out = run_in_main("System.out.println(doubleIt(6));", types);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn final_parameter_in_instance_method() {
    let types = r#"
        static class MathBox {
            int add(final int a, final int b) { return a + b; }
        }
    "#;
    let out = run_in_main(
        "MathBox m = new MathBox(); System.out.println(m.add(3, 4));",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn effectively_final_local_captured_by_lambda() {
    let out = run_main(
        "int seed = 4; java.util.function.IntUnaryOperator inc = n -> n + seed; System.out.println(inc.applyAsInt(1));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn explicit_final_local_captured_by_lambda() {
    let out = run_main(
        "final int seed = 4; java.util.function.IntUnaryOperator inc = n -> n + seed; System.out.println(inc.applyAsInt(1));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn final_field_accessible_from_instance_method() {
    let types = r#"
        static class Reader {
            final int value = 11;
            int read() { return value; }
        }
    "#;
    let out = run_in_main(
        "Reader r = new Reader(); System.out.println(r.read());",
        types,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn final_local_in_for_loop_initializer_style() {
    let out = run_main(
        "int sum = 0; for (int i = 0; i < 3; i++) { final int step = i + 1; sum += step; } System.out.println(sum);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn final_array_reference_binding_elements_mutable() {
    let out = run_main(
        "final int[] data = {1, 2, 3}; data[1] = 9; System.out.println(data[1]);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn final_class_with_constructor_initializing_final_field() {
    let types = r#"
        final class Point {
            final int x;
            Point(int v) { x = v; }
        }
    "#;
    let out = run_in_main(
        "Point p = new Point(8); System.out.println(p.x);",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn final_method_returns_computed_value_from_final_field() {
    let types = r#"
        static class Unit {
            final int factor = 3;
            final int scale(int n) { return n * factor; }
        }
    "#;
    let out = run_in_main(
        "Unit u = new Unit(); System.out.println(u.scale(5));",
        types,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn final_static_field_used_in_instance_method() {
    let types = r#"
        static class Limits {
            static final int CAP = 50;
            int clamp(int n) { return n > CAP ? CAP : n; }
        }
    "#;
    let out = run_in_main(
        "Limits l = new Limits(); System.out.println(l.clamp(80));",
        types,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn final_local_in_nested_block_scope() {
    let out = run_main(
        "int out = 0; { final int inner = 6; out = inner; } System.out.println(out);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn final_field_inheritance_visible_in_subclass() {
    let types = r#"
        static class Parent {
            final int id = 2;
        }
        static class Child extends Parent {}
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c.id);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn final_method_in_grandparent_visible_in_grandchild() {
    let types = r#"
        static class A {
            final int mark() { return 5; }
        }
        static class B extends A {}
        static class C extends B {}
    "#;
    let out = run_in_main(
        "C c = new C(); System.out.println(c.mark());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn final_class_with_to_string_override() {
    let types = r#"
        final class Label {
            final String text = "done";
            public String toString() { return text; }
        }
    "#;
    let out = run_in_main(
        "Label l = new Label(); System.out.println(l);",
        types,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn final_local_boolean_used_in_condition() {
    let out = run_main(
        "final boolean ok = true; if (ok) { System.out.println(\"yes\"); }",
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn final_local_double_precision_value() {
    let out = run_main(
        "final double pi = 3.5; System.out.println(pi + 0.5);",
    );
    assert_eq!(out, vec!["4.0"]);
}

#[test]
fn final_parameter_string_uppercase_helper() {
    let types = r#"
        static String shout(final String s) { return s.toUpperCase(); }
    "#;
    let out = run_in_main(
        r#"System.out.println(shout("go"));"#,
        types,
    );
    assert_eq!(out, vec!["GO"]);
}

#[test]
fn final_static_int_used_in_expression() {
    let types = r#"
        static class Const {
            static final int OFFSET = 3;
        }
    "#;
    let out = run_in_main(
        "System.out.println(10 + Const.OFFSET);",
        types,
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn final_field_multiple_instances_have_independent_values() {
    let types = r#"
        static class Slot {
            final int n;
            Slot(int v) { n = v; }
        }
    "#;
    let out = run_in_main(
        "Slot a = new Slot(1); Slot b = new Slot(2); System.out.println(a.n + b.n);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn final_method_chain_with_second_non_final_method() {
    let types = r#"
        static class Pipe {
            final int start() { return 1; }
            int next(int n) { return n + 1; }
        }
    "#;
    let out = run_in_main(
        "Pipe p = new Pipe(); System.out.println(p.next(p.start()));",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn final_local_char_literal() {
    let out = run_main("final char c = 'K'; System.out.println(c);");
    assert_eq!(out, vec!["K"]);
}

#[test]
fn final_class_static_method_accesses_final_static_field() {
    let types = r#"
        final class Gate {
            static final int CODE = 99;
            static int code() { return CODE; }
        }
    "#;
    let out = run_in_main("System.out.println(Gate.code());", types);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn final_blank_field_assigned_once_in_constructor_body() {
    let types = r#"
        static class Holder {
            final String msg;
            Holder() {
                msg = "ready";
            }
        }
    "#;
    let out = run_in_main(
        "Holder h = new Holder(); System.out.println(h.msg);",
        types,
    );
    assert_eq!(out, vec!["ready"]);
}

#[test]
fn final_local_long_suffix_literal() {
    let out = run_main("final long n = 1_000L; System.out.println(n);");
    assert_eq!(out, vec!["1000"]);
}

#[test]
fn final_method_with_final_parameter_sum() {
    let types = r#"
        static class Adder {
            final int sum(final int a, final int b) { return a + b; }
        }
    "#;
    let out = run_in_main(
        "Adder a = new Adder(); System.out.println(a.sum(4, 5));",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn final_class_cannot_be_extended_compile_semantics_documented_by_usage() {
    let types = r#"
        final class Terminal {
            String ping() { return "ok"; }
        }
    "#;
    let out = run_in_main(
        "Terminal t = new Terminal(); System.out.println(t.ping());",
        types,
    );
    assert_eq!(out, vec!["ok"]);
}
