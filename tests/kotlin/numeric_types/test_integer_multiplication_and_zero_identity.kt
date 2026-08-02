// vybe-test: kotlin/numeric_types/test_integer_multiplication_and_zero_identity
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((7 * 6).toString(), "42")
            __check((7 * 0).toString(), "0")
            __check((0 * 9).toString(), "0")
            __check((-3 * 4).toString(), "-12")
        }
