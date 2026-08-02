// vybe-test: kotlin/numeric_types/test_integer_addition_and_subtraction
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((10 + 5).toString(), "15")
            __check((10 - 5).toString(), "5")
            __check((-3 + 7).toString(), "4")
            __check((3 - 10).toString(), "-7")
        }
