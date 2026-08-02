// vybe-test: kotlin/math_builtins/test_floor_negative_values
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((floor(-3.2)).toString(), "-4")
            __check((floor(-3.9)).toString(), "-4")
            __check((floor(-0.1)).toString(), "-1")
            __check((floor(-2.0)).toString(), "-2")
        }
