// vybe-test: kotlin/when_expressions/test_when_with_guard_first_match_is_used
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun check(value: Int): String {
            return when {
                value > 0 && value % 2 == 0 -> "positive-even"
                value > 0 -> "positive-odd"
                value < 0 -> "negative"
                else -> "zero"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((check(6)).toString(), "positive-even")
            __check((check(5)).toString(), "positive-odd")
            __check((check(-3)).toString(), "negative")
            __check((check(0)).toString(), "zero")
        }
