// vybe-test: kotlin/when_expressions/test_when_used_with_assignment_and_expression_result
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun build(value: Int): String {
            val out = when (value) {
                1 -> "one"
                2 -> "two"
                else -> "other"
            }
            return out
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((build(1)).toString(), "one")
            __check((build(2)).toString(), "two")
            __check((build(9)).toString(), "other")
        }
