kotlin_run_test!(
    test_plus_assign,
    r#"fun main() { var a = 1; a += 2; println(a) }"#,
    &["3"]
);

kotlin_run_test!(
    test_minus_assign,
    r#"fun main() { var a = 10; a -= 3; println(a) }"#,
    &["7"]
);

kotlin_run_test!(
    test_times_assign,
    r#"fun main() { var a = 4; a *= 3; println(a) }"#,
    &["12"]
);

kotlin_run_test!(
    test_div_assign,
    r#"fun main() { var a = 8; a /= 2; println(a) }"#,
    &["4"]
);

kotlin_run_test!(
    test_rem_assign,
    r#"fun main() { var a = 10; a %= 4; println(a) }"#,
    &["2"]
);

kotlin_run_test!(
    test_plus_assign_string,
    r#"fun main() { var s = "a"; s += "b"; println(s) }"#,
    &["ab"]
);

kotlin_run_test!(
    test_minus_assign_long,
    r#"fun main() { var a: Long = 20L; a -= 5L; println(a) }"#,
    &["15"]
);

kotlin_run_test!(
    test_times_assign_float,
    r#"fun main() { var a = 2.5f; a *= 4f; println(a) }"#,
    &["10.0"]
);

kotlin_run_test!(
    test_div_assign_double,
    r#"fun main() { var a = 9.0; a /= 3.0; println(a) }"#,
    &["3.0"]
);

kotlin_run_test!(
    test_rem_assign_int,
    r#"fun main() { var a = 17; a %= 5; println(a) }"#,
    &["2"]
);

kotlin_run_test!(
    test_increment,
    r#"fun main() { var a = 2; a++; println(a) }"#,
    &["3"]
);

kotlin_run_test!(
    test_increment_postfix,
    r#"fun main() { var a = 2; val b = a++; println(a); println(b) }"#,
    &["3", "2"]
);

kotlin_run_test!(
    test_decrement,
    r#"fun main() { var a = 5; a--; println(a) }"#,
    &["4"]
);

kotlin_run_test!(
    test_decrement_prefix,
    r#"fun main() { var a = 8; val b = --a; println(a); println(b) }"#,
    &["7", "7"]
);

kotlin_run_test!(
    test_string_builder_add,
    r#"fun main() {
        val sb = StringBuilder()
        sb.append("a")
        sb.append("b")
        println(sb)
    }"#,
    &["ab"]
);

kotlin_run_test!(
    test_plus_assign_with_expression,
    r#"fun main() { var a = 1; val b = 2; a += b * 2; println(a) }"#,
    &["5"]
);

kotlin_run_test!(
    test_nested_plus_assign,
    r#"fun main() { var a = 1; a += 2; a += 3; println(a) }"#,
    &["6"]
);

kotlin_run_test!(
    test_reassign_string_builder,
    r#"fun main() { var s = "x"; s += "y"; s += "z"; println(s) }"#,
    &["xyz"]
);

kotlin_run_test!(
    test_long_chain_assign,
    r#"fun main() {
        var a: Long = 2L
        a = a + 1
        a *= 2
        a -= 1
        println(a)
    }"#,
    &["5"]
);

kotlin_run_test!(
    test_bitshift_left_assign,
    r#"fun main() { var a = 1; a = a shl 3; println(a) }"#,
    &["8"]
);

kotlin_run_test!(
    test_bitshift_right_assign,
    r#"fun main() { var a = 16; a = a shr 2; println(a) }"#,
    &["4"]
);

kotlin_run_test!(
    test_unsigned_shift_right_not,
    r#"fun main() { var a = -8; a = a ushr 1; println(a) }"#,
    &["2147483644"]
);

kotlin_run_test!(
    test_and_assign,
    r#"fun main() { var a = 6; a = a and 3; println(a) }"#,
    &["2"]
);

kotlin_run_test!(
    test_or_assign,
    r#"fun main() { var a = 2; a = a or 1; println(a) }"#,
    &["3"]
);

kotlin_run_test!(
    test_xor_assign,
    r#"fun main() { var a = 6; a = a xor 3; println(a) }"#,
    &["5"]
);

kotlin_run_test!(
    test_compound_with_negative,
    r#"fun main() { var a = 10; a += -3; println(a) }"#,
    &["7"]
);

kotlin_run_test!(
    test_compound_float_div,
    r#"fun main() { var a = 10.0f; a /= 4f; println(a) }"#,
    &["2.5"]
);

kotlin_run_test!(
    test_compound_double_rem,
    r#"fun main() { var a = 10.0; a = a % 3.0; println(a) }"#,
    &["1.0"]
);

kotlin_run_test!(
    test_assign_inside_function,
    r#"fun bump(v: Int): Int { var x = v; x += 1; return x }
fun main() { println(bump(8)) }"#,
    &["9"]
);

kotlin_run_test!(
    test_assign_with_return_and_print,
    r#"fun main() {
        var a = 1
        println(a)
        a = a + 4
        println(a)
    }"#,
    &["1", "5"]
);

kotlin_run_test!(
    test_assign_loop,
    r#"fun main() {
        var x = 0
        for (i in 1..3) { x += i }
        println(x)
    }"#,
    &["6"]
);

kotlin_run_test!(
    test_assign_while,
    r#"fun main() {
        var a = 1
        var i = 0
        while (i < 4) { a *= 2; i++ }
        println(a)
    }"#,
    &["16"]
);

kotlin_run_test!(
    test_assign_if,
    r#"fun main() {
        var x = 1
        val y = if (x == 1) { x += 5; x } else { x }
        println(y)
    }"#,
    &["6"]
);

kotlin_run_test!(
    test_assign_ternary_like,
    r#"fun main() {
        var x = 2
        val y = if (x < 0) x - 1 else x + 1
        println(y)
    }"#,
    &["3"]
);

kotlin_run_test!(
    test_assign_boolean_not_allowed,
    r#"fun main() {
        var ok = true
        ok = ok && false
        println(ok)
    }"#,
    &["false"]
);

kotlin_run_test!(
    test_assign_with_function,
    r#"fun add(a: Int): Int = a + 10
fun main() {
    var x = 1
    x += add(2)
    println(x)
}"#,
    &["13"]
);

kotlin_run_test!(
    test_assign_chain_two_vars,
    r#"fun main() {
        var a = 1
        var b = 2
        a += b
        b += a
        println(a)
        println(b)
    }"#,
    &["3", "5"]
);

kotlin_run_test!(
    test_assign_mod_nested,
    r#"fun main() {
        var x = 100
        x %= 7
        x += 1
        println(x)
    }"#,
    &["3"]
);
