// vybe-test: kotlin/math_builtins/test_pow_fractional_exponent_small_roots
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(27.0, 1.0 / 3.0)).toString(), "3")
            __check((pow(64.0, 1.0 / 3.0)).toString(), "4")
            __check((pow(4.0, 0.5)).toString(), "2")
            __check((pow(16.0, 0.5)).toString(), "4")
            __check((round(pow(27.0, 1.0 / 3.0))).toString(), "3")
        }
