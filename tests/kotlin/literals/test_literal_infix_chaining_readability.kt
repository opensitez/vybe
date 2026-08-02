// vybe-test: kotlin/literals/test_literal_infix_chaining_readability
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1 + 2 * 3 - 4 / 2
            val grouped = (1 + 2) * (3 - 4 / 2)
            val withFloats = 1 + 2.0 * 2.5 - 3.0
            __check((value).toString(), "5")
            __check((grouped).toString(), "3")
            __check((withFloats).toString(), "3.0")
        }
