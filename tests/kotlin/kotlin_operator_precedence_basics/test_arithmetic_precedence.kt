// vybe-test: kotlin/kotlin_operator_precedence_basics/test_arithmetic_precedence
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_precedence_basics.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 + 2 * 3).toString(), "7")
            __check(((1 + 2) * 3).toString(), "9")
            __check((10 - 2 - 1).toString(), "7")
            __check((10 - (2 - 1)).toString(), "9")
        }
