kotlin_run_test!(
    test_if_expression_value_return,
    r#"
        fun classify(v: Int): String {
            return if (v > 0) "positive" else if (v < 0) "negative" else "zero"
        }

        fun main() {
            println(classify(3))
            println(classify(0))
            println(classify(-1))
        }
    "#,
    &["positive", "zero", "negative"]
);

kotlin_run_test!(
    test_nested_if_expression_blocks,
    r#"
        fun label(v: Int): String {
            return if (v % 2 == 0) {
                if (v > 10) "large-even" else "small-even"
            } else {
                if (v > 10) "large-odd" else "small-odd"
            }
        }

        fun main() {
            println(label(4))
            println(label(9))
            println(label(12))
        }
    "#,
    &["small-even", "small-odd", "large-even"]
);

kotlin_run_test!(
    test_if_with_early_return,
    r#"
        fun classify(v: Int): Int {
            if (v < 0) return -1
            return if (v == 0) 0 else 1
        }

        fun main() {
            println(classify(-3))
            println(classify(0))
            println(classify(2))
        }
    "#,
    &["-1", "0", "1"]
);

kotlin_run_test!(
    test_if_as_statement_changes_outer_variable,
    r#"
        fun main() {
            var total = 0
            if (true) {
                total += 2
            } else {
                total += 9
            }
            if (false) {
                total += 9
            } else if (total == 2) {
                total += 3
            }
            println(total)
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_if_with_boolean_operators,
    r#"
        fun main() {
            val a = 4
            val b = 9
            val result = if (a > 2 && b > 8) "ok" else "bad"
            val second = if (a > 10 || b > 8) "yes" else "no"
            println(result)
            println(second)
        }
    "#,
    &["ok", "yes"]
);

kotlin_run_test!(
    test_if_expression_for_type_aliasing,
    r#"
        fun main() {
            val text: Any = "abc"
            val out = if (text is String) text.length else 0
            println(out)
            val num: Any = 5
            val out2 = if (num is Int) num + 1 else -1
            println(out2)
        }
    "#,
    &["3", "6"]
);

kotlin_run_test!(
    test_if_on_collections,
    r#"
        fun main() {
            val values = listOf(1, 2)
            val size = if (values.isEmpty()) 0 else values.size
            val first = if (values.isNotEmpty()) values[0] else -1
            println(size)
            println(first)
        }
    "#,
    &["2", "1"]
);

kotlin_run_test!(
    test_if_chain_with_elif,
    r#"
        fun classify(v: Int): String {
            return if (v < 0) "neg" else if (v == 0) "zero" else if (v in 1..10) "small" else "big"
        }

        fun main() {
            println(classify(-1))
            println(classify(0))
            println(classify(7))
            println(classify(15))
        }
    "#,
    &["neg", "zero", "small", "big"]
);

kotlin_run_test!(
    test_if_expression_assigning_immutable,
    r#"
        fun main() {
            val a = 7
            val b = if (a > 5) a + 1 else a - 1
            val c = if (a == 10) "ten" else if (a == 7) "seven" else "other"
            println(b)
            println(c)
        }
    "#,
    &["8", "seven"]
);

kotlin_run_test!(
    test_if_with_nullable_else_branch,
    r#"
        fun main() {
            val value: String? = null
            val out = if (value == null) "empty" else value
            println(out)
            val other: String? = "x"
            val out2 = if (other == null) "empty" else other
            println(out2)
        }
    "#,
    &["empty", "x"]
);

kotlin_run_test!(
    test_if_expression_in_loop,
    r#"
        fun main() {
            val values = listOf(1, 2, 3)
            var out = 0
            for (v in values) {
                out += if (v % 2 == 0) 2 else 1
            }
            println(out)
        }
    "#,
    &["5"]
);
