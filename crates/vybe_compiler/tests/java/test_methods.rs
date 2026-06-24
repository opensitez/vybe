use crate::helpers::run_in_main;

#[test]
fn static_method_returns_computed_value() {
    let types = r#"
        static int square(int n) { return n * n; }
    "#;
    let out = run_in_main("System.out.println(square(7));", types);
    assert_eq!(out, vec!["49"]);
}

#[test]
fn recursive_method_computes_factorial() {
    let types = r#"
        static int fact(int n) {
            if (n <= 1) return 1;
            return n * fact(n - 1);
        }
    "#;
    let out = run_in_main("System.out.println(fact(5));", types);
    assert_eq!(out, vec!["120"]);
}

#[test]
fn method_returns_concatenated_string() {
    let types = r#"
        static String greet(String name) { return "hi " + name; }
    "#;
    let out = run_in_main("System.out.println(greet(\"java\"));", types);
    assert_eq!(out, vec!["hi java"]);
}

#[test]
fn overloaded_methods_dispatch_by_parameter_types() {
    let types = r#"
        static String describe(int n) { return "int:" + n; }
        static String describe(double d) { return "dbl:" + d; }
    "#;
    let out = run_in_main(
        "System.out.println(describe(3)); System.out.println(describe(3.14));",
        types,
    );
    assert_eq!(out, vec!["int:3", "dbl:3.14"]);
}

#[test]
fn varargs_method_sums_all_arguments() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sum(1, 2, 3, 4));", types);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn void_method_prints_side_effect() {
    let types = r#"
        static void shout(String msg) { System.out.println(msg.toUpperCase()); }
    "#;
    let out = run_in_main("shout(\"quiet\");", types);
    assert_eq!(out, vec!["QUIET"]);
}
