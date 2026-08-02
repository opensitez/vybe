// vybe-test: kotlin/when_expressions/test_when_with_fallback_expressions_using_arithmetic
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun score(value: Int): String {
            return when (value % 3) {
                0 -> "triple"
                1 -> "plus"
                2 -> "plus2"
                else -> "?"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(10)).toString(), "plus")
            __check((score(11)).toString(), "plus2")
            __check((score(12)).toString(), "triple")
        }
