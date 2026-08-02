// vybe-test: kotlin/kotlin_operator_overflow/test_long_shift_and_overflow
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1L shl 63).toString(), "-9223372036854775808")
            __check(((-1L) ushr 1).toString(), "9223372036854775807")
        }
