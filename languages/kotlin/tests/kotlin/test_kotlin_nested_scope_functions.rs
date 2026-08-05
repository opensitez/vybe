kotlin_run_test!(
    test_local_function_uses_outer_arguments,
    r#"
        fun main() {
            fun format(v: Int): String {
                return "v=$v"
            }
            println(format(7))
        }
    "#,
    &["v=7"]
);

kotlin_run_test!(
    test_local_class_with_accessor,
    r#"
        fun main() {
            class Local(val base: Int) {
                fun value() = base * 2
            }
            val local = Local(4)
            println(local.value())
        }
    "#,
    &["8"]
);

kotlin_run_test!(
    test_nested_lambda_captures_outer_variable,
    r#"
        fun main() {
            var total = 1
            val inc = {
                val local = 2
                total += local
            }
            inc()
            println(total)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_return_from_nested_function,
    r#"
        fun main() {
            fun compute(x: Int): Int {
                fun square(v: Int) = v * v
                return square(x) + 1
            }
            println(compute(4))
        }
    "#,
    &["17"]
);

kotlin_run_test!(
    test_nested_scope_with_if_expression,
    r#"
        fun main() {
            val out = run {
                if (2 > 1) {
                    "yes"
                } else {
                    "no"
                }
            }
            println(out)
        }
    "#,
    &["yes"]
);

kotlin_run_test!(
    test_chain_scoping_functions_locally,
    r#"
        fun main() {
            val out = listOf(1, 2, 3)
                .map { it * 2 }
                .let { numbers -> numbers.filter { it > 2 } }
                .also { println(it.size) }
                .sum()
            println(out)
        }
    "#,
    // Real Kotlin agrees: `[2,4,6].filter { it > 2 }` EXCLUDES 2, so the
    // `also` sees size 2 and the sum is 4+6 = 10.
    &["2", "10"]
);

kotlin_run_test!(
    test_nested_with_and_apply_combo,
    r#"
        class Box {
            var value = 1
            fun bump() { value += 1 }
        }

        fun main() {
            val b = Box().apply {
                bump()
                value += 2
            }.run {
                "$value"
            }
            println(b)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_local_function_with_generic,
    r#"
        fun main() {
            fun <T> firstOf(items: List<T>): T = items[0]
            println(firstOf(listOf("a", "b")))
            println(firstOf(listOf(1, 2, 3)))
        }
    "#,
    &["a", "1"]
);

kotlin_run_test!(
    test_nested_anonymous_object_and_capture,
    r#"
        fun main() {
            fun make(prefix: String) = object {
                fun label(v: Int) = prefix + v
            }

            val p = make("x")
            println(p.label(9))
        }
    "#,
    &["x9"]
);

kotlin_run_test!(
    test_repeated_local_scope_call,
    r#"
        fun main() {
            val values = mutableListOf<Int>()
            run {
                for (i in 1..3) values.add(i)
            }
            println(values.joinToString(""))
        }
    "#,
    &["123"]
);

kotlin_run_test!(
    test_nested_lambda_returns_unit,
    r#"
        fun main() {
            val printer = { text: String ->
                println(text)
                Unit
            }
            printer("ok")
            println("done")
        }
    "#,
    &["ok", "done"]
);

kotlin_run_test!(
    test_scope_function_on_local_result,
    r#"
        fun main() {
            val n = with("ok") {
                length + 1
            }
            println(n)
        }
    "#,
    &["3"]
);
