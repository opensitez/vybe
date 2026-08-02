// vybe-test: kotlin/numeric_types/test_float_division_keeps_fractional_part
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((7.0 / 2.0).toString(), "3.5")
            __check((1.0 / 2.0).toString(), "0.5")
            __check((-7.0 / 2.0).toString(), "-3.5")
        }
