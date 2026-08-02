// vybe-test: kotlin/math_builtins/test_abs_double_negative_and_zero
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(-12.75)).toString(), "12.75")
            __check((abs(0.0)).toString(), "0")
            __check((abs(-0.0)).toString(), "0")
            __check((abs(5.0)).toString(), "5")
        }
