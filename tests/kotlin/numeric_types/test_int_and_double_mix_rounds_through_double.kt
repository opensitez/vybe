// vybe-test: kotlin/numeric_types/test_int_and_double_mix_rounds_through_double
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 5
            __check((value + 2.5).toString(), "7.5")
            __check((value * 1.5).toString(), "7.5")
            __check((value / 2.0).toString(), "2.5")
            __check((10 / 4 + 0.5).toString(), "3.5")
        }
