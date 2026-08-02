// vybe-test: kotlin/math_builtins/test_sqrt_decimal_input
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sqrt(12.25)).toString(), "3.5")
            __check((sqrt(0.81)).toString(), "0.9")
            __check((sqrt(6.25)).toString(), "2.5")
            __check((sqrt(2.56)).toString(), "1.6")
        }
