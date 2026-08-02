// vybe-test: kotlin/when_expressions/test_when_with_guard_condition
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun tag(value: Int): String {
            return when {
                value > 10 -> "gt"
                value == 10 -> "eq"
                else -> "lt"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((tag(12)).toString(), "gt")
            __check((tag(10)).toString(), "eq")
            __check((tag(2)).toString(), "lt")
        }
