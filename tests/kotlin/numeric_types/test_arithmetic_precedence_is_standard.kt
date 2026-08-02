// vybe-test: kotlin/numeric_types/test_arithmetic_precedence_is_standard
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((2 + 3 * 4).toString(), "14")
            __check(((2 + 3) * 4).toString(), "20")
            __check((10 - 6 / 2 + 3).toString(), "10")
            __check((10 - (6 / (2 + 1))).toString(), "8")
        }
