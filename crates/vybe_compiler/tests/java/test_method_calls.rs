use crate::helpers::{run_in_main, run_main};

#[test]
fn static_method_returns_sum_of_two_integers() {
    let types = r#"
        static int add(int a, int b) { return a + b; }
    "#;
    let out = run_in_main("System.out.println(add(12, 30));", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn static_method_returns_boolean_negation() {
    let types = r#"
        static boolean flip(boolean b) { return !b; }
    "#;
    let out = run_in_main("System.out.println(flip(true)); System.out.println(flip(false));", types);
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn static_method_returns_doubled_double() {
    let types = r#"
        static double twice(double d) { return d * 2.0; }
    "#;
    let out = run_in_main("System.out.println(twice(2.5));", types);
    assert_eq!(out, vec!["5.0"]);
}

#[test]
fn static_void_method_prints_uppercased_argument() {
    let types = r#"
        static void announce(String msg) { System.out.println(msg.toUpperCase()); }
    "#;
    let out = run_in_main("announce(\"quiet\");", types);
    assert_eq!(out, vec!["QUIET"]);
}

#[test]
fn instance_method_returns_field_value() {
    let types = r#"
        static class Box {
            int value;
            Box(int v) { value = v; }
            int read() { return value; }
        }
    "#;
    let out = run_in_main("Box b = new Box(17); System.out.println(b.read());", types);
    assert_eq!(out, vec!["17"]);
}

#[test]
fn instance_void_method_mutates_internal_state() {
    let types = r#"
        static class Acc {
            int total = 0;
            void add(int n) { total += n; }
            int get() { return total; }
        }
    "#;
    let out = run_in_main(
        "Acc a = new Acc(); a.add(3); a.add(5); System.out.println(a.get());",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn method_with_single_string_parameter_returns_length() {
    let types = r#"
        static int len(String s) { return s.length(); }
    "#;
    let out = run_in_main("System.out.println(len(\"java\"));", types);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn method_with_three_parameters_computes_average() {
    let types = r#"
        static int avg3(int a, int b, int c) { return (a + b + c) / 3; }
    "#;
    let out = run_in_main("System.out.println(avg3(3, 6, 9));", types);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn zero_argument_static_method_returns_constant() {
    let types = r#"
        static int answer() { return 42; }
    "#;
    let out = run_in_main("System.out.println(answer());", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn recursive_factorial_base_case_and_step() {
    let types = r#"
        static int fact(int n) {
            if (n <= 1) return 1;
            return n * fact(n - 1);
        }
    "#;
    let out = run_in_main("System.out.println(fact(6));", types);
    assert_eq!(out, vec!["720"]);
}

#[test]
fn recursive_fibonacci_returns_expected_sequence_value() {
    let types = r#"
        static int fib(int n) {
            if (n <= 1) return n;
            return fib(n - 1) + fib(n - 2);
        }
    "#;
    let out = run_in_main("System.out.println(fib(10));", types);
    assert_eq!(out, vec!["55"]);
}

#[test]
fn stringbuilder_append_chain_builds_word() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append("j").append("a").append("v").append("a"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["java"]);
}

#[test]
fn stringbuilder_append_then_tostring_returns_full_buffer() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("pre"); sb.append("-"); sb.append("fix"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["pre-fix"]);
}

#[test]
fn stringbuilder_insert_puts_text_at_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ace"); sb.insert(1, "b"); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abce"]);
}

#[test]
fn stringbuilder_reverse_flips_characters() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("stressed"); sb.reverse(); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["desserts"]);
}

#[test]
fn math_abs_called_from_main_returns_positive() {
    let out = run_main("System.out.println(Math.abs(-88));");
    assert_eq!(out, vec!["88"]);
}

#[test]
fn math_max_called_from_main_picks_larger() {
    let out = run_main("System.out.println(Math.max(4, 9));");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn integer_parseint_called_from_main_converts_digits() {
    let out = run_main("System.out.println(Integer.parseInt(\"256\"));");
    assert_eq!(out, vec!["256"]);
}

#[test]
fn method_calls_helper_method_for_computation() {
    let types = r#"
        static int square(int n) { return n * n; }
        static int sumOfSquares(int a, int b) { return square(a) + square(b); }
    "#;
    let out = run_in_main("System.out.println(sumOfSquares(3, 4));", types);
    assert_eq!(out, vec!["25"]);
}

#[test]
fn nested_method_call_expression_evaluates_inner_first() {
    let types = r#"
        static int inc(int n) { return n + 1; }
        static int triple(int n) { return n * 3; }
    "#;
    let out = run_in_main("System.out.println(triple(inc(4)));", types);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn instance_method_returns_this_for_fluent_chaining() {
    let types = r#"
        static class Builder {
            String text = "";
            Builder append(String part) { text = text + part; return this; }
            String build() { return text; }
        }
    "#;
    let out = run_in_main(
        "Builder b = new Builder(); System.out.println(b.append(\"a\").append(\"b\").build());",
        types,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn static_method_invoked_from_within_instance_method() {
    let types = r#"
        static class Util {
            static int doubleIt(int n) { return n * 2; }
            int process(int n) { return doubleIt(n) + 1; }
        }
    "#;
    let out = run_in_main("Util u = new Util(); System.out.println(u.process(5));", types);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn instance_method_called_on_object_created_in_main() {
    let types = r#"
        static class Greeter {
            String greet(String name) { return "hi " + name; }
        }
    "#;
    let out = run_in_main(
        "Greeter g = new Greeter(); System.out.println(g.greet(\"vybe\"));",
        types,
    );
    assert_eq!(out, vec!["hi vybe"]);
}

#[test]
fn method_concatenates_two_strings_with_separator() {
    let types = r#"
        static String join(String a, String b) { return a + ":" + b; }
    "#;
    let out = run_in_main("System.out.println(join(\"foo\", \"bar\"));", types);
    assert_eq!(out, vec!["foo:bar"]);
}

#[test]
fn method_sums_array_elements_passed_as_argument() {
    let types = r#"
        static int total(int[] nums) {
            int sum = 0;
            for (int n : nums) sum += n;
            return sum;
        }
    "#;
    let out = run_in_main("System.out.println(total(new int[] {2, 4, 6}));", types);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn mutually_recursive_methods_detect_even_number() {
    let types = r#"
        static boolean isEven(int n) {
            if (n == 0) return true;
            return isOdd(n - 1);
        }
        static boolean isOdd(int n) {
            if (n == 0) return false;
            return isEven(n - 1);
        }
    "#;
    let out = run_in_main(
        "System.out.println(isEven(4)); System.out.println(isOdd(5));",
        types,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn instance_method_reads_default_initialized_field() {
    let types = r#"
        static class Flag {
            boolean on = true;
            boolean isOn() { return on; }
        }
    "#;
    let out = run_in_main("Flag f = new Flag(); System.out.println(f.isOn());", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn private_static_helper_used_by_public_static_method() {
    let types = r#"
        static int secret(int n) { return n + 10; }
        static int expose(int n) { return secret(n) * 2; }
    "#;
    let out = run_in_main("System.out.println(expose(5));", types);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn overloaded_methods_dispatch_on_int_vs_string() {
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
fn method_returns_null_reference() {
    let types = r#"
        static String nothing() { return null; }
    "#;
    let out = run_in_main("System.out.println(nothing());", types);
    assert_eq!(out, vec!["null"]);
}

#[test]
fn method_accepts_char_parameter_and_returns_code_unit() {
    let types = r#"
        static int code(char c) { return (int) c; }
    "#;
    let out = run_in_main("System.out.println(code('Z'));", types);
    assert_eq!(out, vec!["90"]);
}

#[test]
fn method_accepts_long_parameter_without_overflow_on_small_value() {
    let types = r#"
        static long widen(int n) { return (long) n + 1L; }
    "#;
    let out = run_in_main("System.out.println(widen(100));", types);
    assert_eq!(out, vec!["101"]);
}

#[test]
fn method_returns_float_quotient() {
    let types = r#"
        static float ratio(int a, int b) { return (float) a / (float) b; }
    "#;
    let out = run_in_main("System.out.println(ratio(7, 2));", types);
    assert_eq!(out, vec!["3.5"]);
}

#[test]
fn method_conditional_return_picks_branch() {
    let types = r#"
        static String sign(int n) {
            if (n > 0) return "pos";
            if (n < 0) return "neg";
            return "zero";
        }
    "#;
    let out = run_in_main(
        "System.out.println(sign(3)); System.out.println(sign(-2)); System.out.println(sign(0));",
        types,
    );
    assert_eq!(out, vec!["pos", "neg", "zero"]);
}

#[test]
fn method_loop_accumulates_running_total() {
    let types = r#"
        static int sumTo(int n) {
            int total = 0;
            for (int i = 1; i <= n; i++) total += i;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sumTo(5));", types);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn stringbuilder_delete_removes_subrange() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcdef"); sb.delete(1, 4); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["adf"]);
}

#[test]
fn stringbuilder_delete_char_at_removes_single_index() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("abcde"); sb.deleteCharAt(2); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["abde"]);
}

#[test]
fn stringbuilder_chained_append_insert_reverse() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("ab"); sb.append("cd").insert(2, "-").reverse(); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["dc-ba"]);
}

#[test]
fn constructor_calls_instance_method_via_this() {
    let types = r#"
        static class Init {
            int value;
            Init(int seed) { value = bump(seed); }
            int bump(int n) { return n + 1; }
        }
    "#;
    let out = run_in_main("Init i = new Init(9); System.out.println(i.value);", types);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn static_initializer_value_read_by_static_method() {
    let types = r#"
        static int base = 5;
        static int offset(int n) { return base + n; }
    "#;
    let out = run_in_main("System.out.println(offset(7));", types);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn main_body_calls_top_level_static_method_directly() {
    let types = r#"
        static int triple(int n) { return n * 3; }
    "#;
    let out = run_in_main("System.out.println(triple(8));", types);
    assert_eq!(out, vec!["24"]);
}

#[test]
fn deeply_nested_static_method_calls_evaluate_correctly() {
    let types = r#"
        static int a(int n) { return n + 1; }
        static int b(int n) { return a(n) + 1; }
        static int c(int n) { return b(n) + 1; }
    "#;
    let out = run_in_main("System.out.println(c(1));", types);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn method_does_not_mutate_primitive_parameter_copy() {
    let types = r#"
        static int bumpCopy(int n) { n = n + 5; return n; }
    "#;
    let out = run_in_main("int x = 3; System.out.println(bumpCopy(x)); System.out.println(x);", types);
    assert_eq!(out, vec!["8", "3"]);
}

#[test]
fn instance_method_reads_static_field_updated_by_prior_call() {
    let types = r#"
        static class Tally {
            static int hits = 0;
            void mark() { hits++; }
            int count() { return hits; }
        }
    "#;
    let out = run_in_main(
        "Tally t1 = new Tally(); Tally t2 = new Tally(); t1.mark(); t2.mark(); t2.mark(); System.out.println(t2.count());",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn empty_stringbuilder_tostring_returns_empty_string() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); System.out.println(sb.toString().length());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_append_int_coerces_to_text() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder(); sb.append(42); System.out.println(sb.toString());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn method_early_return_skips_remaining_logic() {
    let types = r#"
        static int firstPositive(int a, int b) {
            if (a > 0) return a;
            if (b > 0) return b;
            return 0;
        }
    "#;
    let out = run_in_main("System.out.println(firstPositive(-1, 9));", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn recursive_gcd_computes_greatest_common_divisor() {
    let types = r#"
        static int gcd(int a, int b) {
            if (b == 0) return a;
            return gcd(b, a % b);
        }
    "#;
    let out = run_in_main("System.out.println(gcd(48, 18));", types);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn recursive_power_computes_exponentiation() {
    let types = r#"
        static int pow(int base, int exp) {
            if (exp == 0) return 1;
            return base * pow(base, exp - 1);
        }
    "#;
    let out = run_in_main("System.out.println(pow(2, 8));", types);
    assert_eq!(out, vec!["256"]);
}

#[test]
fn stringbuilder_length_reports_character_count() {
    let out = run_main(
        r#"StringBuilder sb = new StringBuilder("hello"); System.out.println(sb.length());"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn instance_method_invoked_on_field_access_expression() {
    let types = r#"
        static class Pair {
            String left;
            String right;
            Pair(String l, String r) { left = l; right = r; }
            String combine() { return left + right; }
        }
    "#;
    let out = run_in_main(
        "Pair p = new Pair(\"foo\", \"bar\"); System.out.println(p.combine());",
        types,
    );
    assert_eq!(out, vec!["foobar"]);
}
