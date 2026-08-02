// vybe-test: kotlin/kotlin_return_expressions/test_when_as_expression_with_multiple_cases
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun describe(v: Int): String = when (v) {
            in 1..3 -> "small"
            in 4..9 -> "mid"
            else -> "other"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(2)).toString(), "small")
            __check((describe(8)).toString(), "mid")
            __check((describe(20)).toString(), "other")
        }
