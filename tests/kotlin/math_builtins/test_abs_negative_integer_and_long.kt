// vybe-test: kotlin/math_builtins/test_abs_negative_integer_and_long
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(-12)).toString(), "12")
            __check((abs(-12345)).toString(), "12345")
            __check((abs(-123L)).toString(), "123")
            __check((abs(-999_999_999L)).toString(), "999999999")
        }
