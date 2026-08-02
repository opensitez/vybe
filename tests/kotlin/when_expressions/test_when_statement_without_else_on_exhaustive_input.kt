// vybe-test: kotlin/when_expressions/test_when_statement_without_else_on_exhaustive_input
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun colorCode(color: String): Int {
            return when (color) {
                "red" -> 1
                "green" -> 2
                "blue" -> 3
                else -> 0
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((colorCode("red")).toString(), "1")
            __check((colorCode("green")).toString(), "2")
            __check((colorCode("blue")).toString(), "3")
            __check((colorCode("black")).toString(), "0")
        }
