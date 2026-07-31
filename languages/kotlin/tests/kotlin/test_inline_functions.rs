kotlin_run_test!(
    test_inline_function_runs_lambda_once,
    r#"
        inline fun once(block: () -> Int): Int = block()

        fun main() {
            println(once { 5 + 2 })
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_inline_function_with_receiver,
    r#"
        inline fun String.wrap(prefix: String, suffix: String): String = prefix + this + suffix

        fun main() {
            println("k".wrap("<", ">"))
        }
    "#,
    &["<k>"]
);

kotlin_run_test!(
    test_inline_function_accumulator,
    r#"
        inline fun <T> fold(start: T, values: List<T>, op: (T, T) -> T): T {
            var out = start
            for (value in values) {
                out = op(out, value)
            }
            return out
        }

        fun main() {
            println(fold(0, listOf(1, 2, 3), { a, b -> a + b }))
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_inline_higher_order_transform,
    r#"
        inline fun <T, R> mapOrNull(value: T, transform: (T) -> R?): R? = transform(value)

        fun main() {
            println(mapOrNull(4) { if (it > 2) it * 2 else null })
        }
    "#,
    &["8"]
);

kotlin_run_test!(
    test_inline_void_style,
    r#"
        inline fun tap(value: Int, action: (Int) -> Unit): Int {
            action(value)
            return value
        }

        fun main() {
            var seen = 0
            val out = tap(3) { v -> seen += v }
            println(out)
            println(seen)
        }
    "#,
    &["3", "3"]
);

kotlin_run_test!(
    test_inline_chaining_calls,
    r#"
        inline fun firstNonEmpty(values: List<String>): String {
            for (value in values) {
                if (value.isNotEmpty()) return value
            }
            return ""
        }

        fun main() {
            println(firstNonEmpty(listOf("", "kotlin", "x")))
        }
    "#,
    &["kotlin"]
);

kotlin_run_test!(
    test_inline_boolean_mapper,
    r#"
        inline fun <T> allMatch(values: List<T>, check: (T) -> Boolean): Boolean {
            for (value in values) {
                if (!check(value)) return false
            }
            return true
        }

        fun main() {
            println(allMatch(listOf(2, 4, 6)) { it % 2 == 0 })
            println(allMatch(listOf(2, 3, 6)) { it % 2 == 0 })
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_inline_string_mapper,
    r#"
        inline fun format(prefix: String, value: Int, fmt: (Int) -> String): String {
            return prefix + fmt(value)
        }

        fun main() {
            println(format("v=", 4) { v -> v.toString() })
        }
    "#,
    &["v=4"]
);

kotlin_run_test!(
    test_inline_generic_identity,
    r#"
        inline fun <T> identity(value: T): T = value

        fun main() {
            println(identity("kotlin"))
            println(identity(9))
        }
    "#,
    &["kotlin", "9"]
);

kotlin_run_test!(
    test_inline_sequence_step,
    r#"
        inline fun <T> firstOrFallback(values: List<T>, fallback: T): T {
            for (item in values) {
                return item
            }
            return fallback
        }

        fun main() {
            println(firstOrFallback(listOf(8, 9), 0))
            println(firstOrFallback(emptyList(), 3))
        }
    "#,
    &["8", "3"]
);

kotlin_run_test!(
    test_inline_predicate_chain,
    r#"
        inline fun check(value: Int, tests: (Int) -> Boolean, onFail: () -> String): String {
            return if (tests(value)) "ok" else onFail()
        }

        fun main() {
            println(check(5, { it > 2 }) { "bad" })
            println(check(1, { it > 2 }) { "bad" })
        }
    "#,
    &["ok", "bad"]
);
