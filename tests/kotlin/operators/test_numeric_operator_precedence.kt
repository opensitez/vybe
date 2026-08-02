// vybe-test: kotlin/operators/test_numeric_operator_precedence
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 + 2 * 3).toString(), "7")
            __check(((1 + 2) * 3).toString(), "9")
            __check((10 - 3 * 2).toString(), "4")
            __check((10 / 2 + 1).toString(), "6")
        }
