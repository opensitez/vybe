kotlin_run_test!(
    test_let_returns_last_expression,
    r#"
        fun describe(v: Int): String {
            return v.let { it + 1 }
                .let { it * 2 }
                .toString()
        }

        fun main() {
            println(describe(3))
        }
    "#,
    &["8"]
);

kotlin_run_test!(
    test_apply_mutates_receiver,
    r#"
        data class Box(var value: Int)

        fun main() {
            val box = Box(1).apply { value += 3 }
            println(box.value)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_also_preserves_receiver_and_side_effects,
    r#"
        fun main() {
            val out = mutableListOf<Int>()
            val values = intArrayOf(1, 2, 3).toMutableList().also {
                out.addAll(it)
                it.add(4)
            }
            println(values.joinToString(","))
            println(out.joinToString(","))
        }
    "#,
    &["1,2,3,4", "1,2,3"]
);

kotlin_run_test!(
    test_run_and_with_return_different_receiver,
    r#"
        fun main() {
            val withValue = with("a") { this + "bc" }
            val runValue = "a".run { uppercase() + "bc" }
            println(withValue)
            println(runValue)
        }
    "#,
    &["abc", "Abc"]
);

kotlin_run_test!(
    test_take_if_keeps_matching,
    r#"
        fun main() {
            val a = 7.takeIf { it > 5 }
            val b = 3.takeIf { it > 5 }
            println(a)
            println(b)
        }
    "#,
    &["7", "null"]
);

kotlin_run_test!(
    test_take_unless_filters_when_match,
    r#"
        fun main() {
            val a = 7.takeUnless { it > 5 }
            val b = 3.takeUnless { it > 5 }
            println(a)
            println(b)
        }
    "#,
    &["null", "3"]
);

kotlin_run_test!(
    test_chained_scoping_functions,
    r#"
        fun main() {
            val text = "kotlin"
                .also { println(it) }
                .let { it.reversed() }
                .run { toUpperCase() }
            println(text)
        }
    "#,
    &["kotlin", "NILTOK"]
);

kotlin_run_test!(
    test_scoping_function_on_nullable,
    r#"
        fun main() {
            val value: String? = "abc"
            val a = value?.let { it + "d" }
            val b = value?.let { null }
            println(a)
            println(b)
        }
    "#,
    &["abcd", "null"]
);

kotlin_run_test!(
    test_with_configures_receiver_without_reference,
    r#"
        data class Accumulator(var total: Int)

        fun main() {
            val acc = Accumulator(0).apply {
                total += 1
                total += 2
            }
            val done = with(acc) {
                total += 5
                total
            }
            println(acc.total)
            println(done)
        }
    "#,
    &["8", "8"]
);

kotlin_run_test!(
    test_run_with_non_trivial_return_type,
    r#"
        fun main() {
            val values = listOf(1, 2, 3)
            val score = values.run {
                filter { it % 2 == 1 }.sum()
            }
            println(score)
        }
    "#,
    &["4"]
);
