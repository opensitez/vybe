use crate::helpers::run_in_main;

#[test]
fn overload_dispatches_int_over_double_for_whole_number() {
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
fn overload_dispatches_string_over_int_by_argument_type() {
    let types = r#"
        static String tag(int n) { return "i" + n; }
        static String tag(String s) { return "s" + s; }
    "#;
    let out = run_in_main(
        "System.out.println(tag(7)); System.out.println(tag(\"x\"));",
        types,
    );
    assert_eq!(out, vec!["i7", "sx"]);
}

#[test]
fn overload_selects_one_parameter_version() {
    let types = r#"
        static int size(int a) { return 1; }
        static int size(int a, int b) { return 2; }
    "#;
    let out = run_in_main(
        "System.out.println(size(4)); System.out.println(size(4, 5));",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn overload_selects_two_vs_three_parameter_versions() {
    let types = r#"
        static int arity(int a, int b) { return 2; }
        static int arity(int a, int b, int c) { return 3; }
    "#;
    let out = run_in_main(
        "System.out.println(arity(1, 2)); System.out.println(arity(1, 2, 3));",
        types,
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn overload_int_add_differs_from_double_add() {
    let types = r#"
        static int add(int a, int b) { return a + b; }
        static double add(double a, double b) { return a + b; }
    "#;
    let out = run_in_main(
        "System.out.println(add(2, 3)); System.out.println(add(2.5, 0.5));",
        types,
    );
    assert_eq!(out, vec!["5", "3.0"]);
}

#[test]
fn instance_overload_dispatches_int_and_string() {
    let types = r#"
        static class Printer {
            String show(int n) { return "n" + n; }
            String show(String s) { return "s" + s; }
        }
    "#;
    let out = run_in_main(
        "Printer p = new Printer(); System.out.println(p.show(4)); System.out.println(p.show(\"z\"));",
        types,
    );
    assert_eq!(out, vec!["n4", "sz"]);
}

#[test]
fn overload_boolean_and_int_are_distinct() {
    let types = r#"
        static String code(boolean b) { return b ? "T" : "F"; }
        static String code(int n) { return "I" + n; }
    "#;
    let out = run_in_main(
        "System.out.println(code(true)); System.out.println(code(2));",
        types,
    );
    assert_eq!(out, vec!["T", "I2"]);
}

#[test]
fn overload_char_and_int_select_different_methods() {
    let types = r#"
        static String mark(char c) { return "c" + c; }
        static String mark(int n) { return "i" + n; }
    "#;
    let out = run_in_main(
        "System.out.println(mark('A')); System.out.println(mark(65));",
        types,
    );
    assert_eq!(out, vec!["cA", "i65"]);
}

#[test]
fn overload_zero_args_vs_one_arg() {
    let types = r#"
        static int mode() { return 0; }
        static int mode(int n) { return n; }
    "#;
    let out = run_in_main(
        "System.out.println(mode()); System.out.println(mode(9));",
        types,
    );
    assert_eq!(out, vec!["0", "9"]);
}

#[test]
fn overload_same_count_different_second_parameter_type() {
    let types = r#"
        static String pair(int a, int b) { return "ii" + a + b; }
        static String pair(int a, String b) { return "is" + a + b; }
    "#;
    let out = run_in_main(
        "System.out.println(pair(1, 2)); System.out.println(pair(1, \"x\"));",
        types,
    );
    assert_eq!(out, vec!["ii12", "is1x"]);
}

#[test]
fn overload_float_and_double_are_distinct_targets() {
    let types = r#"
        static String kind(float f) { return "f"; }
        static String kind(double d) { return "d"; }
    "#;
    let out = run_in_main(
        "System.out.println(kind(1.0f)); System.out.println(kind(1.0));",
        types,
    );
    assert_eq!(out, vec!["f", "d"]);
}

#[test]
fn overload_void_print_int_vs_string() {
    let types = r#"
        static void print(int n) { System.out.println(\"i\" + n); }
        static void print(String s) { System.out.println(\"s\" + s); }
    "#;
    let out = run_in_main("print(3); print(\"ok\");", types);
    assert_eq!(out, vec!["i3", "sok"]);
}

#[test]
fn overload_pick_max_for_int_and_double_pairs() {
    let types = r#"
        static int max(int a, int b) { return a >= b ? a : b; }
        static double max(double a, double b) { return a >= b ? a : b; }
    "#;
    let out = run_in_main(
        "System.out.println(max(2, 9)); System.out.println(max(2.5, 1.5));",
        types,
    );
    assert_eq!(out, vec!["9", "2.5"]);
}

#[test]
fn overload_format_one_int_vs_two_ints() {
    let types = r#"
        static String format(int a) { return \"(\" + a + \")\"; }
        static String format(int a, int b) { return \"[\" + a + \",\" + b + \"]\"; }
    "#;
    let out = run_in_main(
        "System.out.println(format(4)); System.out.println(format(4, 5));",
        types,
    );
    assert_eq!(out, vec!["(4)", "[4,5]"]);
}

#[test]
fn overload_three_versions_by_parameter_count() {
    let types = r#"
        static int pack(int a) { return 1; }
        static int pack(int a, int b) { return 2; }
        static int pack(int a, int b, int c) { return 3; }
    "#;
    let out = run_in_main(
        "System.out.println(pack(1)); System.out.println(pack(1, 2)); System.out.println(pack(1, 2, 3));",
        types,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn overload_object_and_string_reference_types() {
    let types = r#"
        static String label(Object o) { return "o"; }
        static String label(String s) { return "s" + s; }
    "#;
    let out = run_in_main(
        "System.out.println(label(\"hi\"));",
        types,
    );
    assert_eq!(out, vec!["shi"]);
}

#[test]
fn overload_instance_chain_calls_two_signatures() {
    let types = r#"
        static class Calc {
            int step(int n) { return n + 1; }
            int step(int n, int m) { return n + m; }
        }
    "#;
    let out = run_in_main(
        "Calc c = new Calc(); System.out.println(c.step(2)); System.out.println(c.step(2, 3));",
        types,
    );
    assert_eq!(out, vec!["3", "5"]);
}

#[test]
fn overload_int_and_long_primitive_forms() {
    let types = r#"
        static String wide(int n) { return "i"; }
        static String wide(long n) { return "l"; }
    "#;
    let out = run_in_main(
        "System.out.println(wide(5)); System.out.println(wide(5L));",
        types,
    );
    assert_eq!(out, vec!["i", "l"]);
}

#[test]
fn overload_string_builder_vs_string_by_type() {
    let types = r#"
        static int len(String s) { return s.length(); }
        static int len(String a, String b) { return a.length() + b.length(); }
    "#;
    let out = run_in_main(
        "System.out.println(len(\"ab\")); System.out.println(len(\"a\", \"bbb\"));",
        types,
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn overload_first_parameter_type_breaks_tie_on_second() {
    let types = r#"
        static String mix(int a, String b) { return "is"; }
        static String mix(String a, int b) { return "si"; }
    "#;
    let out = run_in_main(
        "System.out.println(mix(1, \"x\")); System.out.println(mix(\"x\", 1));",
        types,
    );
    assert_eq!(out, vec!["is", "si"]);
}

#[test]
fn overload_returns_different_string_prefixes() {
    let types = r#"
        static String kind(int n) { return "int"; }
        static String kind(double d) { return "dbl"; }
        static String kind(String s) { return "str"; }
    "#;
    let out = run_in_main(
        "System.out.println(kind(1)); System.out.println(kind(1.0)); System.out.println(kind(\"a\"));",
        types,
    );
    assert_eq!(out, vec!["int", "dbl", "str"]);
}

#[test]
fn overload_two_ints_vs_int_and_string() {
    let types = r#"
        static int score(int a, int b) { return a + b; }
        static int score(int a, String b) { return a + b.length(); }
    "#;
    let out = run_in_main(
        "System.out.println(score(2, 3)); System.out.println(score(2, \"xy\"));",
        types,
    );
    assert_eq!(out, vec!["5", "4"]);
}

#[test]
fn overload_static_methods_share_name_with_instance_methods() {
    let types = r#"
        static class Tool {
            static int work(int n) { return n * 2; }
            int work(int n, int m) { return n + m; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Tool.work(4)); Tool t = new Tool(); System.out.println(t.work(4, 5));",
        types,
    );
    assert_eq!(out, vec!["8", "9"]);
}

#[test]
fn overload_byte_and_short_literals_use_int_fallback() {
    let types = r#"
        static String tiny(byte b) { return "b"; }
        static String tiny(short s) { return "s"; }
        static String tiny(int n) { return "i"; }
    "#;
    let out = run_in_main(
        "System.out.println(tiny((byte) 2)); System.out.println(tiny((short) 2));",
        types,
    );
    assert_eq!(out, vec!["b", "s"]);
}

#[test]
fn overload_three_param_types_permute_second_and_third() {
    let types = r#"
        static String order(int a, int b, int c) { return "iii"; }
        static String order(int a, int b, String c) { return "iis"; }
    "#;
    let out = run_in_main(
        "System.out.println(order(1, 2, 3)); System.out.println(order(1, 2, \"z\"));",
        types,
    );
    assert_eq!(out, vec!["iii", "iis"]);
}
