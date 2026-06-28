use crate::helpers::{run_in_main, run_main};

#[test]
fn anonymous_runnable_executes_run_method() {
    let out = run_main(
        "Runnable r = new Runnable() { public void run() { System.out.println(\"run\"); } }; r.run();",
    );
    assert_eq!(out, vec!["run"]);
}

#[test]
fn anonymous_runnable_prints_custom_message() {
    let out = run_main(
        "Runnable r = new Runnable() { public void run() { System.out.println(\"hello\"); } }; r.run();",
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn anonymous_comparator_sorts_ascending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(3); list.add(1); list.add(2); list.sort(new java.util.Comparator<Integer>() { public int compare(Integer a, Integer b) { return a - b; } }); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn anonymous_comparator_sorts_descending() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(3); list.add(2); list.sort(new java.util.Comparator<Integer>() { public int compare(Integer a, Integer b) { return b - a; } }); System.out.println(list.get(0)); System.out.println(list.get(2));",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn anonymous_interface_single_method() {
    let types = r#"
        static interface Greeter { String greet(); }
    "#;
    let out = run_in_main(
        "Greeter g = new Greeter() { public String greet() { return \"hi\"; } }; System.out.println(g.greet());",
        types,
    );
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn anonymous_interface_multiple_methods() {
    let types = r#"
        static interface Ops {
            int add(int a, int b);
            int mul(int a, int b);
        }
    "#;
    let out = run_in_main(
        "Ops ops = new Ops() { public int add(int a, int b) { return a + b; } public int mul(int a, int b) { return a * b; } }; System.out.println(ops.add(2, 3)); System.out.println(ops.mul(2, 3));",
        types,
    );
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn anonymous_abstract_class_extends_and_implements_method() {
    let types = r#"
        static abstract class Base { abstract int value(); }
    "#;
    let out = run_in_main(
        "Base b = new Base() { int value() { return 42; } }; System.out.println(b.value());",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn anonymous_class_with_instance_field() {
    let types = r#"
        static interface Counter { int next(); }
    "#;
    let out = run_in_main(
        "Counter c = new Counter() { int n = 0; public int next() { n++; return n; } }; System.out.println(c.next()); System.out.println(c.next());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn anonymous_class_captures_final_local() {
    let out = run_main(
        "int base = 10; Runnable r = new Runnable() { public void run() { System.out.println(base + 1); } }; r.run();",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn anonymous_class_assigned_to_interface_reference() {
    let types = r#"
        static interface Named { String name(); }
    "#;
    let out = run_in_main(
        "Named n = new Named() { public String name() { return \"anon\"; } }; System.out.println(n.name());",
        types,
    );
    assert_eq!(out, vec!["anon"]);
}

#[test]
fn anonymous_class_assigned_to_abstract_base() {
    let types = r#"
        static abstract class Node { abstract int depth(); }
    "#;
    let out = run_in_main(
        "Node n = new Node() { int depth() { return 2; } }; System.out.println(n.depth());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn anonymous_class_in_array_of_interfaces() {
    let types = r#"
        static interface Valued { int value(); }
    "#;
    let out = run_in_main(
        "Valued[] arr = new Valued[] { new Valued() { public int value() { return 1; } }, new Valued() { public int value() { return 2; } } }; System.out.println(arr[0].value()); System.out.println(arr[1].value());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn anonymous_class_returned_from_factory_method() {
    let types = r#"
        static interface Box { int get(); }
        static class Factory {
            Box make(int n) {
                return new Box() { public int get() { return n; } };
            }
        }
    "#;
    let out = run_in_main(
        "Factory f = new Factory(); System.out.println(f.make(7).get());",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn anonymous_class_overrides_single_method() {
    let types = r#"
        static class Base { String label() { return "base"; } }
    "#;
    let out = run_in_main(
        "Base b = new Base() { String label() { return \"anon\"; } }; System.out.println(b.label());",
        types,
    );
    assert_eq!(out, vec!["anon"]);
}

#[test]
fn anonymous_class_calls_super_method() {
    let types = r#"
        static class Base { int step() { return 1; } }
    "#;
    let out = run_in_main(
        "Base b = new Base() { int step() { return super.step() + 4; } }; System.out.println(b.step());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn anonymous_class_instanceof_interface_true() {
    let types = r#"
        static interface Marker {}
    "#;
    let out = run_in_main(
        "Marker m = new Marker() {}; System.out.println(m instanceof Marker);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn anonymous_class_instanceof_own_superclass() {
    let types = r#"
        static class Parent {}
    "#;
    let out = run_in_main(
        "Parent p = new Parent() {}; System.out.println(p instanceof Parent);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn anonymous_comparator_in_collections_sort() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"b\"); list.add(\"a\"); java.util.Collections.sort(list, new java.util.Comparator<String>() { public int compare(String a, String b) { return a.compareTo(b); } }); System.out.println(list.get(0));",
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn anonymous_class_with_void_method_side_effect() {
    let types = r#"
        static interface Logger { void log(String msg); }
    "#;
    let out = run_in_main(
        "Logger log = new Logger() { public void log(String msg) { System.out.println(msg); } }; log.log(\"logged\");",
        types,
    );
    assert_eq!(out, vec!["logged"]);
}

#[test]
fn anonymous_class_overrides_to_string() {
    let out = run_main(
        "Object o = new Object() { public String toString() { return \"anon\"; } }; System.out.println(o.toString());",
    );
    assert_eq!(out, vec!["anon"]);
}

#[test]
fn anonymous_class_multiple_instances_distinct() {
    let types = r#"
        static interface Id { int id(); }
    "#;
    let out = run_in_main(
        "Id a = new Id() { public int id() { return 1; } }; Id b = new Id() { public int id() { return 2; } }; System.out.println(a.id()); System.out.println(b.id());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn anonymous_class_extends_concrete_parent_overrides() {
    let types = r#"
        static class Animal { String sound() { return "..."; } }
    "#;
    let out = run_in_main(
        "Animal a = new Animal() { String sound() { return \"meow\"; } }; System.out.println(a.sound());",
        types,
    );
    assert_eq!(out, vec!["meow"]);
}

#[test]
fn anonymous_class_adds_method_beyond_abstract_parent() {
    let types = r#"
        static abstract class Base { abstract int core(); }
    "#;
    let out = run_in_main(
        "Base b = new Base() { int core() { return 2; } int extra() { return core() + 3; } }; System.out.println(b.core());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn anonymous_class_as_method_argument() {
    let types = r#"
        static interface Task { void run(); }
        static class Runner { void execute(Task t) { t.run(); } }
    "#;
    let out = run_in_main(
        "Runner r = new Runner(); r.execute(new Task() { public void run() { System.out.println(\"done\"); } });",
        types,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn anonymous_class_with_boolean_logic() {
    let types = r#"
        static interface Check { boolean ok(int n); }
    "#;
    let out = run_in_main(
        "Check c = new Check() { public boolean ok(int n) { return n > 0; } }; System.out.println(c.ok(1)); System.out.println(c.ok(-1));",
        types,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn anonymous_class_with_string_concatenation() {
    let types = r#"
        static interface Builder { String build(String part); }
    "#;
    let out = run_in_main(
        "Builder b = new Builder() { public String build(String part) { return \"pre-\" + part; } }; System.out.println(b.build(\"fix\"));",
        types,
    );
    assert_eq!(out, vec!["pre-fix"]);
}

#[test]
fn anonymous_class_chain_of_two_wrappers() {
    let types = r#"
        static interface IntFn { int apply(int n); }
    "#;
    let out = run_in_main(
        "IntFn step1 = new IntFn() { public int apply(int n) { return n + 1; } }; IntFn step2 = new IntFn() { public int apply(int n) { return step1.apply(n) * 2; } }; System.out.println(step2.apply(3));",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn anonymous_class_implements_nested_interface() {
    let types = r#"
        static class Outer {
            static interface InnerApi { int code(); }
        }
    "#;
    let out = run_in_main(
        "Outer.InnerApi api = new Outer.InnerApi() { public int code() { return 9; } }; System.out.println(api.code());",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn anonymous_class_accesses_enclosing_static_field() {
    let types = r#"
        static class Host { static int seed = 5; }
    "#;
    let out = run_in_main(
        "Runnable r = new Runnable() { public void run() { System.out.println(Host.seed); } }; r.run();",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn anonymous_class_counter_increments_state() {
    let types = r#"
        static interface Tally { int bump(); int read(); }
    "#;
    let out = run_in_main(
        "Tally t = new Tally() { int count = 0; public int bump() { count++; return count; } public int read() { return count; } }; System.out.println(t.bump()); System.out.println(t.read());",
        types,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn anonymous_class_double_dispatch_via_interface() {
    let types = r#"
        static interface Shape { int area(); }
    "#;
    let out = run_in_main(
        "Shape s1 = new Shape() { public int area() { return 4; } }; Shape s2 = new Shape() { public int area() { return 9; } }; System.out.println(s1.area() + s2.area());",
        types,
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn anonymous_class_nested_in_method() {
    let types = r#"
        static class Maker {
            Runnable make() {
                return new Runnable() { public void run() { System.out.println("nested"); } };
            }
        }
    "#;
    let out = run_in_main("Maker m = new Maker(); m.make().run();", types);
    assert_eq!(out, vec!["nested"]);
}

#[test]
fn anonymous_class_with_super_constructor_call() {
    let types = r#"
        static class Base {
            int value;
            Base(int value) { this.value = value; }
        }
    "#;
    let out = run_in_main(
        "Base b = new Base(6) { int doubled() { return value * 2; } }; System.out.println(b.value);",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn anonymous_class_used_in_foreach_list() {
    let types = r#"
        static interface Label { String text(); }
    "#;
    let out = run_in_main(
        "java.util.ArrayList<Label> list = new java.util.ArrayList<Label>(); list.add(new Label() { public String text() { return \"a\"; } }); list.add(new Label() { public String text() { return \"b\"; } }); System.out.println(list.get(0).text()); System.out.println(list.get(1).text());",
        types,
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn anonymous_class_implements_generic_comparator_for_strings() {
    let out = run_main(
        "java.util.ArrayList<String> list = new java.util.ArrayList<String>(); list.add(\"z\"); list.add(\"m\"); list.sort(new java.util.Comparator<String>() { public int compare(String a, String b) { return a.compareTo(b); } }); System.out.println(list.get(0)); System.out.println(list.get(1));",
    );
    assert_eq!(out, vec!["m", "z"]);
}

#[test]
fn anonymous_class_with_two_implemented_methods_on_interface() {
    let types = r#"
        static interface Pair {
            int left();
            int right();
        }
    "#;
    let out = run_in_main(
        "Pair p = new Pair() { public int left() { return 1; } public int right() { return 2; } }; System.out.println(p.left() + p.right());",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn anonymous_runnable_assigned_then_reassigned() {
    let out = run_main(
        "Runnable r = new Runnable() { public void run() { System.out.println(\"first\"); } }; r.run(); r = new Runnable() { public void run() { System.out.println(\"second\"); } }; r.run();",
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn anonymous_class_extends_and_implements_interface() {
    let types = r#"
        static class Base { String prefix() { return "x"; } }
        static interface Suffix { String suffix(); }
    "#;
    let out = run_in_main(
        "Suffix s = new Base() implements Suffix { public String suffix() { return prefix() + \"-y\"; } }; System.out.println(s.suffix());",
        types,
    );
    assert_eq!(out, vec!["x-y"]);
}

#[test]
fn anonymous_class_with_numeric_return_from_interface() {
    let types = r#"
        static interface Calc { double compute(double x); }
    "#;
    let out = run_in_main(
        "Calc c = new Calc() { public double compute(double x) { return x * x; } }; System.out.println(c.compute(3.0));",
        types,
    );
    assert_eq!(out, vec!["9.0"]);
}

#[test]
fn anonymous_class_in_switch_like_conditional() {
    let types = r#"
        static interface Mode { String label(); }
    "#;
    let out = run_in_main(
        "Mode mode = new Mode() { public String label() { return \"fast\"; } }; String out = mode.label(); System.out.println(out);",
        types,
    );
    assert_eq!(out, vec!["fast"]);
}
