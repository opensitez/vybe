// vybe-test: kotlin/math_builtins/test_floor_for_positive_and_zero_values
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((floor(3.0)).toString(), "3")
            __check((floor(3.9)).toString(), "3")
            __check((floor(0.1)).toString(), "0")
            __check((floor(-0.0)).toString(), "0")
        }
