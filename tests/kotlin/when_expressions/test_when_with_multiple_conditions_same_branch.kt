// vybe-test: kotlin/when_expressions/test_when_with_multiple_conditions_same_branch
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun classify(value: Int): String {
            return when (value) {
                1, 2, 3 -> "low"
                4, 5, 6 -> "mid"
                else -> "high"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(2)).toString(), "low")
            __check((classify(6)).toString(), "mid")
            __check((classify(9)).toString(), "high")
        }
