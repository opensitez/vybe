// vybe-test: kotlin/math_builtins/test_ceil_negative_values
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((ceil(-3.2)).toString(), "-3")
            __check((ceil(-3.9)).toString(), "-3")
            __check((ceil(-0.9)).toString(), "0")
            __check((ceil(-2.0)).toString(), "-2")
        }
