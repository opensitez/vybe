kotlin_run_test!(
    test_unary_plus,
    r#"
        fun main() {
            val x = +5
            println(x)
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_unary_minus,
    r#"
        fun main() {
            val x = -5
            println(x)
        }
    "#,
    &["-5"]
);

kotlin_run_test!(
    test_unary_not_true,
    r#"
        fun main() {
            println(!true)
            println(!false)
        }
    "#,
    &["false", "true"]
);

kotlin_run_test!(
    test_prefix_increment,
    r#"
        fun main() {
            var x = 1
            println(++x)
            println(x)
        }
    "#,
    &["2", "2"]
);

kotlin_run_test!(
    test_postfix_increment,
    r#"
        fun main() {
            var x = 1
            println(x++)
            println(x)
        }
    "#,
    &["1", "2"]
);

kotlin_run_test!(
    test_prefix_decrement,
    r#"
        fun main() {
            var x = 3
            println(--x)
            println(x)
        }
    "#,
    &["2", "2"]
);

kotlin_run_test!(
    test_postfix_decrement,
    r#"
        fun main() {
            var x = 3
            println(x--)
            println(x)
        }
    "#,
    &["3", "2"]
);

kotlin_run_test!(
    test_unary_on_various_math,
    r#"
        fun main() {
            val a = -3 + +2
            val b = -(-3)
            println(a)
            println(b)
        }
    "#,
    &["-1", "3"]
);

kotlin_run_test!(
    test_unary_boolean_chain,
    r#"
        fun main() {
            val a = true
            val b = false
            println(!(a && b))
            println(!(a || b))
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_unary_expression_with_nullable,
    r#"
        fun main() {
            val value: Int? = 5
            println(value ?: -1)
            println(value?.let { -it } ?: -1)
        }
    "#,
    &["5", "-5"]
);

kotlin_run_test!(
    test_unary_with_long,
    r#"
        fun main() {
            val x = 10L
            println(-x)
            println(+x)
        }
    "#,
    &["-10", "10"]
);

kotlin_run_test!(
    test_unary_with_double,
    r#"
        fun main() {
            val x = 1.5
            println(-x)
            println(+x)
        }
    "#,
    &["-1.5", "1.5"]
);

kotlin_run_test!(
    test_unary_after_computation,
    r#"
        fun main() {
            var out = 0
            var x = 2
            out += +x
            out += -x
            println(out)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_not_chain,
    r#"
        fun main() {
            println(!!false)
            println(!true)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_increment_in_loop,
    r#"
        fun main() {
            var x = 0
            var out = 0
            repeat(3) {
                out += ++x
            }
            println(out)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_decrement_in_sum,
    r#"
        fun main() {
            var x = 10
            var out = 0
            out += x--
            out += x
            println(out)
        }
    "#,
    &["19"]
);

kotlin_run_test!(
    test_mix_increments,
    r#"
        fun main() {
            var x = 1
            var out = 0
            out += x++
            out += ++x
            println(out)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_unary_negate_char_code,
    r#"
        fun main() {
            val c = 'a'.code
            println(c)
            println(-c)
        }
    "#,
    &["97", "-97"]
);

kotlin_run_test!(
    test_unary_boolean_algebra,
    r#"
        fun main() {
            val a = true
            val b = false
            val out = ! (a || b) == false
            println(out)
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_unary_nested,
    r#"
        fun main() {
            val x = -(-(-2))
            println(x)
        }
    "#,
    &["-2"]
);

kotlin_run_test!(
    test_unary_in_expression,
    r#"
        fun sign(v: Int): Int {
            return if (v > 0) +1 else -1
        }
        fun main() {
            println(sign(3))
            println(sign(-2))
        }
    "#,
    &["1", "-1"]
);

kotlin_run_test!(
    test_unary_preserves_equality,
    r#"
        fun main() {
            val a = -(-3)
            val b = 3
            println(a == b)
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_unary_sequence_with_if,
    r#"
        fun main() {
            val x = -3
            println(if (x < 0) -x else x)
            println(if (x > 0) -x else +x)
        }
    "#,
    &["3", "-3"]
);

kotlin_run_test!(
    test_unary_on_function_call,
    r#"
        fun value(): Int = 4
        fun main() {
            println(-value())
            println(+value())
        }
    "#,
    &["-4", "4"]
);

kotlin_run_test!(
    test_unary_with_overflow,
    r#"
        fun main() {
            val x = Int.MIN_VALUE
            val y = -x
            println(y)
        }
    "#,
    &["-2147483648"]
);

kotlin_run_test!(
    test_unary_negate_zero,
    r#"
        fun main() {
            val x = -0
            println(x)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_unary_apply_chain,
    r#"
        fun main() {
            val a = -(+1)
            val b = (+(-2))
            println(a)
            println(b)
        }
    "#,
    &["-1", "-2"]
);

kotlin_run_test!(
    test_boolean_negation_chain,
    r#"
        fun main() {
            val a = !(!true)
            val b = !!false
            println(a)
            println(b)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_increment_on_expression,
    r#"
        fun main() {
            var x = 1
            val y = ++x + x++
            println(y)
            println(x)
        }
    "#,
    &["3", "3"]
);

kotlin_run_test!(
    test_decrement_in_conditions,
    r#"
        fun main() {
            var x = 4
            val out = if (--x > 2) "gt" else "lte"
            println(out)
        }
    "#,
    &["gt"]
);
