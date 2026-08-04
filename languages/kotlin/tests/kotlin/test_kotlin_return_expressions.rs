kotlin_run_test!(
    test_if_as_expression_true_path,
    r#"
        fun classify(v: Int): String {
            return if (v > 0) "positive" else "non-positive"
        }

        fun main() {
            println(classify(1))
            println(classify(0))
        }
    "#,
    &["positive", "non-positive"]
);

kotlin_run_test!(
    test_when_as_expression_with_multiple_cases,
    r#"
        fun describe(v: Int): String = when (v) {
            in 1..3 -> "small"
            in 4..9 -> "mid"
            else -> "other"
        }

        fun main() {
            println(describe(2))
            println(describe(8))
            println(describe(20))
        }
    "#,
    &["small", "mid", "other"]
);

kotlin_run_test!(
    test_expression_function_syntax,
    r#"
        fun square(x: Int): Int = x * x

        fun main() {
            println(square(4))
        }
    "#,
    &["16"]
);

kotlin_run_test!(
    test_try_as_expression,
    r#"
        fun safe(v: Int): Int = try {
            val inv = 10 / v
            inv
        } catch (e: ArithmeticException) {
            0
        }

        fun main() {
            println(safe(2))
            println(safe(0))
        }
    "#,
    &["5", "0"]
);

kotlin_run_test!(
    test_early_return_in_expression_lambda,
    r#"
        fun runWithGuard(v: Int): String {
            return run {
                if (v == 0) return@run "zero"
                "value"
            }
        }

        fun main() {
            println(runWithGuard(0))
            println(runWithGuard(2))
        }
    "#,
    &["zero", "value"]
);

kotlin_run_test!(
    test_returning_result_from_nested_block,
    r#"
        fun compute(v: Int): Int {
            val out = run {
                if (v < 5) {
                    return@run v * 2
                }
                v
            }
            return out
        }

        fun main() {
            println(compute(3))
            println(compute(7))
        }
    "#,
    &["6", "7"]
);

kotlin_run_test!(
    test_ternary_like_usage_with_if,
    r#"
        fun main() {
            val value = if (10 % 2 == 0) "even" else "odd"
            println(value)
        }
    "#,
    &["even"]
);

kotlin_run_test!(
    test_return_from_function_with_named_arguments,
    r#"
        fun join(prefix: String, value: Int): String {
            return "$prefix$value"
        }

        fun main() {
            println(join(prefix = "x", value = 9))
        }
    "#,
    &["x9"]
);

kotlin_run_test!(
    test_return_last_statement_in_block,
    r#"
        fun score(v: Int): Int {
            return {
                if (v > 1) {
                    v + 1
                } else {
                    v
                }
            }()
        }

        fun main() {
            println(score(2))
            println(score(0))
        }
    "#,
    &["3", "0"]
);

kotlin_run_test!(
    test_conditional_return_to_label,
    r#"
        fun main() {
            val out = run outer@{
                val text = "x"
                if (text.isEmpty()) return@outer "no"
                "yes"
            }
            println(out)
        }
    "#,
    &["yes"]
);

kotlin_run_test!(
    test_try_expression_value_with_finally,
    r#"
        fun main() {
            val out = try {
                "ok"
            } finally {
                println("final")
            }
            println(out)
        }
    "#,
    &["final", "ok"]
);
