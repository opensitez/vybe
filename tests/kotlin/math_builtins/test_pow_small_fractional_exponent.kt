// vybe-test: kotlin/math_builtins/test_pow_small_fractional_exponent
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(16.0, 0.5)).toString(), "4")
            __check((pow(16.0, 0.25)).toString(), "2")
            __check((pow(1.0, 999.0)).toString(), "1")
            __check((pow(9.0, 0.5)).toString(), "3")
        }
