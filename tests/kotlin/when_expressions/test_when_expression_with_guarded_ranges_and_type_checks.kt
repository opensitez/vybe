// vybe-test: kotlin/when_expressions/test_when_expression_with_guarded_ranges_and_type_checks
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun describe(value: Any): String {
            return when {
                value is Int && value < 0 -> "negative"
                value is Int && value == 0 -> "zero"
                value is Int -> "positive"
                value is String && value.isNotBlank() -> "word"
                value == null -> "none"
                else -> "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(-1)).toString(), "negative")
            __check((describe(0)).toString(), "zero")
            __check((describe(5)).toString(), "positive")
            __check((describe("a")).toString(), "word")
            __check((describe(null)).toString(), "none")
        }
