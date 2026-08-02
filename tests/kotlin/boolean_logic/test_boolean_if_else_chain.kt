// vybe-test: kotlin/boolean_logic/test_boolean_if_else_chain
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 7
            val status = if (value > 10) false
                         else if (value > 5 && value < 10) true
                         else false
            __check((status).toString(), "true")
        }
