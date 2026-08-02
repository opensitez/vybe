// vybe-test: kotlin/math_builtins/test_abs_and_round_combo
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(round(-2.6) + pow(2.0, 3.0))).toString(), "11")
            __check((abs(round(3.4) - pow(3.0, 2.0))).toString(), "1")
            __check((abs(pow(2.0, 2.0) - pow(2.0, 3.0) + round(0.5))).toString(), "3")
        }
