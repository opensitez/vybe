// vybe-test: kotlin/math_builtins/test_round_small_integers
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(0.0)).toString(), "0")
            __check((round(-0.0)).toString(), "0")
            __check((round(0.49)).toString(), "0")
            __check((round(-0.49)).toString(), "0")
            __check((round(999.6)).toString(), "1000")
            __check((round(-999.6)).toString(), "-1000")
        }
