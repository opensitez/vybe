// vybe-test: kotlin/math_builtins/test_abs_zero_and_positive_integers
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(0)).toString(), "0")
            __check((abs(12)).toString(), "12")
            __check((abs(999)).toString(), "999")
        }
