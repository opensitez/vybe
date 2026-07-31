kotlin_run_test!(
    test_capture_mutable_counter,
    r#"
        fun main() {
            var count = 0
            val inc = { count += 1 }
            inc()
            inc()
            println(count)
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_capture_var_with_default_argument,
    r#"
        fun main() {
            var text = "a"
            val append = { suffix: String -> text += suffix }
            append("b")
            append("c")
            println(text)
        }
    "#,
    &["abc"]
);

kotlin_run_test!(
    test_capture_array_mutation_in_lambda,
    r#"
        fun main() {
            val list = IntArray(3)
            val mutate = { idx: Int, value: Int -> list[idx] = value }
            mutate(0, 5)
            mutate(1, 6)
            println(list.joinToString(","))
        }
    "#,
    &["5,6,0"]
);

kotlin_run_test!(
    test_capture_function_reference,
    r#"
        fun make(prefix: String): (Int) -> String {
            return { v -> prefix + v.toString() }
        }

        fun main() {
            val f = make("x")
            println(f(7))
        }
    "#,
    &["x7"]
);

kotlin_run_test!(
    test_nested_lambda_shadowing,
    r#"
        fun main() {
            var value = 1
            val outer = {
                val value = 10
                { value + 1 }
            }
            println(outer()())
            println(value)
        }
    "#,
    &["11", "1"]
);

kotlin_run_test!(
    test_capture_in_map_and_for_each,
    r#"
        fun main() {
            var total = 0
            listOf(1, 2, 3).forEach { total += it }
            println(total)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_capture_in_reduce_function,
    r#"
        fun main() {
            var marker = ""
            val values = listOf("a", "b")
            values.reduce { acc, item ->
                marker = acc + item
                marker
            }
            println(marker)
        }
    "#,
    &["ab"]
);

kotlin_run_test!(
    test_capture_data_class_property,
    r#"
        data class State(var value: Int)

        fun main() {
            val state = State(1)
            val bump = { state.value++ }
            bump()
            bump()
            println(state.value)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_capture_boolean_toggle,
    r#"
        fun main() {
            var on = false
            val toggle = { on = !on }
            toggle()
            toggle()
            println(on)
        }
    "#,
    &["false"]
);

kotlin_run_test!(
    test_capture_and_return_factory,
    r#"
        fun makeAdder(base: Int): (Int) -> Int {
            return { x -> x + base }
        }

        fun main() {
            val add5 = makeAdder(5)
            println(add5(10))
        }
    "#,
    &["15"]
);
