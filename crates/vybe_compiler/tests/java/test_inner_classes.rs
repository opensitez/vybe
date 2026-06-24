use crate::helpers::{run_in_main, run_main};

#[test]
fn static_nested_class_instantiated_with_qualified_name() {
    let types = r#"
        static class Outer {
            static class Inner { int value = 6; }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(); System.out.println(inner.value);",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn static_nested_method_returns_constant() {
    let types = r#"
        static class Outer {
            static class Inner { int code() { return 42; } }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(); System.out.println(inner.code());",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn static_nested_with_parameterized_constructor() {
    let types = r#"
        static class Outer {
            static class Inner {
                int value;
                Inner(int value) { this.value = value; }
            }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(9); System.out.println(inner.value);",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn doubly_nested_static_class_accessible() {
    let types = r#"
        static class A {
            static class B {
                static class C { int depth = 3; }
            }
        }
    "#;
    let out = run_in_main(
        "A.B.C c = new A.B.C(); System.out.println(c.depth);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn static_nested_static_field_shared_across_instances() {
    let types = r#"
        static class Outer {
            static class Inner { static int count = 0; static void bump() { count++; } }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner.bump(); Outer.Inner.bump(); System.out.println(Outer.Inner.count);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn static_nested_static_method_without_outer_instance() {
    let types = r#"
        static class Outer {
            static class Inner { static int doubleIt(int n) { return n * 2; } }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Outer.Inner.doubleIt(5));",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn outer_static_field_visible_to_static_nested() {
    let types = r#"
        static class Outer {
            static int shared = 100;
            static class Inner { int read() { return shared; } }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(); System.out.println(inner.read());",
        types,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn static_nested_modifies_outer_static_field() {
    let types = r#"
        static class Outer {
            static int tally = 0;
            static class Inner { void add(int n) { tally = tally + n; } }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(); inner.add(7); System.out.println(Outer.tally);",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn static_nested_field_defaults_to_zero() {
    let types = r#"
        static class Outer {
            static class Inner { int unset; }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(); System.out.println(inner.unset);",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn static_nested_overrides_to_string() {
    let types = r#"
        static class Outer {
            static class Inner {
                String toString() { return "inner"; }
            }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(); System.out.println(inner.toString());",
        types,
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn static_nested_used_as_field_type_in_outer() {
    let types = r#"
        static class Outer {
            static class Inner { int value = 4; }
            Inner holder = new Inner();
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.holder.value);",
        types,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn static_nested_implements_interface_method() {
    let types = r#"
        static interface Named { String name(); }
        static class Outer {
            static class Inner implements Named { String name() { return "inner"; } }
        }
    "#;
    let out = run_in_main(
        "Named n = new Outer.Inner(); System.out.println(n.name());",
        types,
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn static_nested_same_name_in_different_outers() {
    let types = r#"
        static class Alpha { static class Inner { int tag() { return 1; } } }
        static class Beta { static class Inner { int tag() { return 2; } } }
    "#;
    let out = run_in_main(
        "System.out.println(new Alpha.Inner().tag()); System.out.println(new Beta.Inner().tag());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn instance_inner_reads_enclosing_field() {
    let types = r#"
        static class Outer {
            int value = 42;
            class Inner { int read() { return value; } }
            Inner create() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.create().read());",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn instance_inner_created_via_outer_factory_method() {
    let types = r#"
        static class Outer {
            int seed = 5;
            class Inner { int doubled() { return seed * 2; } }
            Inner spawn() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.spawn().doubled());",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn instance_inner_method_returns_outer_field_sum() {
    let types = r#"
        static class Outer {
            int a = 3;
            int b = 4;
            class Inner { int sum() { return a + b; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.make().sum());",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn instance_inner_accesses_outer_this_explicitly() {
    let types = r#"
        static class Outer {
            int value = 11;
            class Inner { int viaOuter() { return Outer.this.value; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.make().viaOuter());",
        types,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn multiple_instance_inners_from_same_outer_share_state() {
    let types = r#"
        static class Outer {
            int counter = 0;
            class Inner { int read() { return counter; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); outer.counter = 9; System.out.println(outer.make().read());",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn instance_inner_calls_outer_instance_method() {
    let types = r#"
        static class Outer {
            int base() { return 2; }
            class Inner { int doubled() { return base() * 2; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.make().doubled());",
        types,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn instance_inner_each_outer_has_distinct_inner_state() {
    let types = r#"
        static class Outer {
            int id;
            Outer(int id) { this.id = id; }
            class Inner { int ownerId() { return id; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer a = new Outer(1); Outer b = new Outer(2); System.out.println(a.make().ownerId()); System.out.println(b.make().ownerId());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn inner_reads_modified_outer_field_after_mutation() {
    let types = r#"
        static class Outer {
            int value = 1;
            class Inner { int read() { return value; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); outer.value = 20; System.out.println(outer.make().read());",
        types,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn instance_inner_shadows_outer_field_name_via_this() {
    let types = r#"
        static class Outer {
            int value = 1;
            class Inner {
                int value = 2;
                int outerValue() { return Outer.this.value; }
            }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.make().outerValue());",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn outer_exposes_inner_type_reference() {
    let types = r#"
        static class Outer {
            class Inner { String label() { return "nested"; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.make().label());",
        types,
    );
    assert_eq!(out, vec!["nested"]);
}

#[test]
fn local_class_in_method_returns_computed_value() {
    let types = r#"
        static class Util {
            int compute() {
                class Local { int value = 7; }
                Local loc = new Local();
                return loc.value;
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.compute());",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn local_class_with_instance_method() {
    let types = r#"
        static class Util {
            String greet(String name) {
                class Local {
                    String hello(String who) { return "hi-" + who; }
                }
                Local loc = new Local();
                return loc.hello(name);
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.greet(\"java\"));",
        types,
    );
    assert_eq!(out, vec!["hi-java"]);
}

#[test]
fn local_class_captures_method_local_variable() {
    let types = r#"
        static class Util {
            int scale(int base) {
                int factor = 3;
                class Local { int apply() { return base * factor; } }
                Local loc = new Local();
                return loc.apply();
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.scale(4));",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn local_class_in_loop_accumulates_sum() {
    let types = r#"
        static class Util {
            int sumThree() {
                int total = 0;
                for (int i = 1; i <= 3; i++) {
                    class Local { int value; Local(int v) { value = v; } }
                    Local loc = new Local(i);
                    total = total + loc.value;
                }
                return total;
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.sumThree());",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn local_class_in_main_method_body() {
    let out = run_main(
        "class Local { int value = 5; } Local loc = new Local(); System.out.println(loc.value);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn local_class_declared_inside_if_branch() {
    let types = r#"
        static class Util {
            int pick(boolean flag) {
                if (flag) {
                    class Local { int value = 1; }
                    return new Local().value;
                } else {
                    class Local { int value = 2; }
                    return new Local().value;
                }
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.pick(true)); System.out.println(util.pick(false));",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn local_class_inside_try_block() {
    let types = r#"
        static class Util {
            int safe() {
                try {
                    class Local { int value = 8; }
                    return new Local().value;
                } catch (RuntimeException e) {
                    return -1;
                }
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.safe());",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn local_class_extends_outer_static_class() {
    let types = r#"
        static class Base { int baseValue() { return 1; } }
        static class Util {
            int derived() {
                class Local extends Base { int total() { return baseValue() + 4; } }
                return new Local().total();
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.derived());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn static_nested_with_string_field() {
    let types = r#"
        static class Outer {
            static class Inner { String text = "nested"; }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(); System.out.println(inner.text);",
        types,
    );
    assert_eq!(out, vec!["nested"]);
}

#[test]
fn instance_inner_with_string_from_outer() {
    let types = r#"
        static class Outer {
            String prefix = "pre";
            class Inner { String label() { return prefix + "-inner"; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.make().label());",
        types,
    );
    assert_eq!(out, vec!["pre-inner"]);
}

#[test]
fn static_nested_multiple_instances_independent_fields() {
    let types = r#"
        static class Outer {
            static class Inner { int value; }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner a = new Outer.Inner(); Outer.Inner b = new Outer.Inner(); a.value = 1; b.value = 2; System.out.println(a.value); System.out.println(b.value);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn local_class_with_constructor_args() {
    let types = r#"
        static class Util {
            int read(int seed) {
                class Local {
                    int value;
                    Local(int v) { value = v; }
                }
                return new Local(seed).value;
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.read(15));",
        types,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn static_nested_accesses_outer_instance_field_via_outer_ref() {
    let types = r#"
        static class Outer {
            int outerValue = 6;
            static class Inner {
                int read(Outer o) { return o.outerValue; }
            }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); Outer.Inner inner = new Outer.Inner(); System.out.println(inner.read(outer));",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn instance_inner_boolean_field_from_outer() {
    let types = r#"
        static class Outer {
            boolean active = true;
            class Inner { boolean isActive() { return active; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.make().isActive());",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_class_in_while_loop_counts_iterations() {
    let types = r#"
        static class Util {
            int countTo(int limit) {
                int i = 0;
                int total = 0;
                while (i < limit) {
                    class Local { int step; Local(int s) { step = s; } }
                    total = total + new Local(i + 1).step;
                    i++;
                }
                return total;
            }
        }
    "#;
    let out = run_in_main(
        "Util util = new Util(); System.out.println(util.countTo(3));",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn static_nested_chain_of_two_instances() {
    let types = r#"
        static class Outer {
            static class Inner {
                int next() { return 3; }
            }
            Inner first() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); System.out.println(outer.first().next());",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn instance_inner_returns_outer_field_after_increment() {
    let types = r#"
        static class Outer {
            int count = 0;
            void bump() { count++; }
            class Inner { int read() { return count; } }
            Inner make() { return new Inner(); }
        }
    "#;
    let out = run_in_main(
        "Outer outer = new Outer(); outer.bump(); outer.bump(); System.out.println(outer.make().read());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}
