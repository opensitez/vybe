kotlin_run_cases! {
    test_expression_function => (r#"
        fun double(v: Int): Int = v * 2
        fun label(v: Int) = "v" + v.toString()

        fun main() {
            println(double(3))
            println(label(4))
        }
    "#, vec!["6", "v4"]),
    test_expression_property => (r#"
        class Box {
            val value: Int get() = 1 + 2
        }

        fun main() {
            println(Box().value)
        }
    "#, vec!["3"]),
    test_expression_lambda => (r#"
        fun main() {
            val sum = { a: Int, b: Int -> a + b }
            println(sum(4, 5).toString())
        }
    "#, vec!["9"]),
}
