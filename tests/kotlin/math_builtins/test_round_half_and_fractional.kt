// vybe-test: kotlin/math_builtins/test_round_half_and_fractional
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(2.4)).toString(), "2")
            __check((round(2.5)).toString(), "3")
            __check((round(2.6)).toString(), "3")
            __check((round(-2.4)).toString(), "-2")
            __check((round(-2.5)).toString(), "-3")
            __check((round(-2.6)).toString(), "-3")
        }
