// vybe-test: kotlin/boolean_logic/test_boolean_in_when_conditions
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 2
            val result = when {
                value % 2 == 0 && value > 1 -> "even"
                value < 0 && value % 2 == 0 -> "neg"
                else -> "other"
            }
            __check((result).toString(), "even")
        }
