kotlin_run_cases! {
    test_if_as_expression => (r#"
        fun classify(value: Int): String {
            return if (value > 0) {
                "pos"
            } else {
                "non"
            }
        }

        fun main() {
            println(classify(1))
            println(classify(0))
        }
    "#, vec!["pos", "non"]),
    test_when_as_expression => (r#"
        fun whenResult(value: Int): String {
            return when (value) {
                1 -> "one"
                2 -> "two"
                else -> "many"
            }
        }

        fun main() {
            println(whenResult(2))
            println(whenResult(4))
        }
    "#, vec!["two", "many"]),
    test_block_scope_value => (r#"
        fun main() {
            val result = {
                val x = 2
                val y = 3
                x + y
            }
            println(result)
        }
    "#, vec!["5"]),
}
