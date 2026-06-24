use crate::helpers::{run_in_main, run_main};

#[test]
fn static_int_field_defaults_to_zero() {
    let types = r#"
        static class Stats {
            static int total;
        }
    "#;
    let out = run_in_main("System.out.println(Stats.total);", types);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn static_int_field_explicit_initializer() {
    let types = r#"
        static class Config {
            static int port = 3000;
        }
    "#;
    let out = run_in_main("System.out.println(Config.port);", types);
    assert_eq!(out, vec!["3000"]);
}

#[test]
fn static_string_literal_initializer() {
    let types = r#"
        static class Registry {
            static String name = "default";
        }
    "#;
    let out = run_in_main("System.out.println(Registry.name);", types);
    assert_eq!(out, vec!["default"]);
}

#[test]
fn static_method_adds_two_integers() {
    let types = r#"
        static class Math2 {
            static int add(int a, int b) { return a + b; }
        }
    "#;
    let out = run_in_main("System.out.println(Math2.add(10, 7));", types);
    assert_eq!(out, vec!["17"]);
}

#[test]
fn static_method_zero_args_returns_constant() {
    let types = r#"
        static class Const {
            static int answer() { return 42; }
        }
    "#;
    let out = run_in_main("System.out.println(Const.answer());", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn static_method_callable_before_any_instance_created() {
    let types = r#"
        static class Early {
            static int ping() { return 1; }
            Early() {}
        }
    "#;
    let out = run_in_main("System.out.println(Early.ping());", types);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn static_method_mutates_static_field() {
    let types = r#"
        static class Counter {
            static int n = 0;
            static void bump() { n++; }
        }
    "#;
    let out = run_in_main("Counter.bump(); Counter.bump(); System.out.println(Counter.n);", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn instance_method_increments_static_counter() {
    let types = r#"
        static class Widget {
            static int created = 0;
            Widget() { created++; }
        }
    "#;
    let out = run_in_main(
        "new Widget(); new Widget(); System.out.println(Widget.created);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn static_field_shared_across_all_instances() {
    let types = r#"
        static class Shared {
            static int hits = 0;
            void touch() { hits++; }
        }
    "#;
    let out = run_in_main(
        "Shared a = new Shared(); Shared b = new Shared(); a.touch(); b.touch(); b.touch(); System.out.println(Shared.hits);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn static_boolean_field_initializer() {
    let types = r#"
        static class Toggle {
            static boolean enabled = true;
        }
    "#;
    let out = run_in_main("System.out.println(Toggle.enabled);", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn static_long_field_initializer() {
    let types = r#"
        static class Limits {
            static long max = 1_000_000L;
        }
    "#;
    let out = run_in_main("System.out.println(Limits.max);", types);
    assert_eq!(out, vec!["1000000"]);
}

#[test]
fn static_method_calls_sibling_static_method() {
    let types = r#"
        static class Chain {
            static int doubleIt(int n) { return n + n; }
            static int quad(int n) { return doubleIt(doubleIt(n)); }
        }
    "#;
    let out = run_in_main("System.out.println(Chain.quad(3));", types);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn static_field_read_from_instance_method_via_class_name() {
    let types = r#"
        static class Reader {
            static int value = 9;
            int read() { return Reader.value; }
        }
    "#;
    let out = run_in_main("Reader r = new Reader(); System.out.println(r.read());", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn static_counter_incremented_from_main_body() {
    let types = r#"
        static class Tally {
            static int count = 0;
        }
    "#;
    let out = run_in_main(
        "Tally.count++; Tally.count++; System.out.println(Tally.count);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn static_init_expression_with_arithmetic() {
    let types = r#"
        static class Expr {
            static int sum = 3 + 4;
        }
    "#;
    let out = run_in_main("System.out.println(Expr.sum);", types);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn static_init_expression_uses_math_abs() {
    let types = r#"
        static class AbsVal {
            static int n = (int) Math.abs(-12);
        }
    "#;
    let out = run_in_main("System.out.println(AbsVal.n);", types);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn static_method_qualified_invocation_from_nested_scope() {
    let types = r#"
        static class Outer {
            static int id() { return 5; }
        }
    "#;
    let out = run_in_main(
        "int x = 0; { x = Outer.id(); } System.out.println(x);",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn static_field_updated_after_instance_creation() {
    let types = r#"
        static class Mutable {
            static int value = 1;
        }
    "#;
    let out = run_in_main(
        "new Mutable(); Mutable.value = 8; System.out.println(Mutable.value);",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn static_method_returns_current_static_field() {
    let types = r#"
        static class Store {
            static int slot = 11;
            static int get() { return slot; }
        }
    "#;
    let out = run_in_main("System.out.println(Store.get());", types);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn static_field_init_runs_once_for_all_reads() {
    let types = r#"
        static class Once {
            static int seed = 7;
        }
    "#;
    let out = run_in_main(
        "System.out.println(Once.seed); System.out.println(Once.seed);",
        types,
    );
    assert_eq!(out, vec!["7", "7"]);
}

#[test]
fn static_method_with_three_parameters() {
    let types = r#"
        static class Sum3 {
            static int add(int a, int b, int c) { return a + b + c; }
        }
    "#;
    let out = run_in_main("System.out.println(Sum3.add(1, 2, 3));", types);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn static_double_field_initializer() {
    let types = r#"
        static class Rates {
            static double ratio = 2.5;
        }
    "#;
    let out = run_in_main("System.out.println(Rates.ratio);", types);
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn static_block_sets_field_on_class_load() {
    let types = r#"
        static class Loader {
            static int ready;
            static { ready = 42; }
        }
    "#;
    let out = run_in_main("System.out.println(Loader.ready);", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn static_block_combines_with_field_initializer() {
    let types = r#"
        static class Mix {
            static int base = 1;
            static int total;
            static { total = base + 4; }
        }
    "#;
    let out = run_in_main("System.out.println(Mix.total);", types);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn multiple_static_blocks_apply_in_sequence() {
    let types = r#"
        static class Ordered {
            static int step;
            static { step = 1; }
            static { step = step + 2; }
        }
    "#;
    let out = run_in_main("System.out.println(Ordered.step);", types);
    assert_eq!(out, vec!["3"]);
}
