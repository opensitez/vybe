// vybe-test: kotlin/literals/test_parenthesized_and_unary_literals
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((-(1 + 2)).toString(), "-3")
            __check((-(2 * 3)).toString(), "-6")
            __check((+5).toString(), "5")
            __check((+(-5)).toString(), "-5")
            __check(((-1L) + 2L).toString(), "1")
        }
