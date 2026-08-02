// vybe-test: kotlin/kotlin_operator_overflow/test_signed_shift_right_with_negative
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((-4) shr 1).toString(), "-2")
            __check(((-4) ushr 1).toString(), "2147483646")
        }
