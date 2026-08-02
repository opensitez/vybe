// vybe-test: kotlin/math_builtins/test_sqrt_negative_is_not_a_number
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sqrt(-4.0).isNaN()).toString(), "true")
            __check((sqrt(-1.5).isNaN()).toString(), "true")
        }
