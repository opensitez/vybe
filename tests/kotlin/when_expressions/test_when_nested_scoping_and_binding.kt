// vybe-test: kotlin/when_expressions/test_when_nested_scoping_and_binding
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun describe(a: Int, b: Int): String {
            return when (a) {
                0 -> when (b) {
                    0 -> "a0b0"
                    else -> "a0bN"
                }
                else -> when {
                    b == 0 -> "aNb0"
                    b > 10 -> "aNbH"
                    else -> "aNbL"
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(0, 0)).toString(), "a0b0")
            __check((describe(0, 4)).toString(), "a0bN")
            __check((describe(5, 12)).toString(), "aNbH")
        }
