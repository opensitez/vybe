// vybe-test: kotlin/bitwise_operations/test_long_shift_right_arithmetic
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Long = 64L
            val negative: Long = -64L
            __check((value shr 2).toString(), "16")
            __check((negative shr 3).toString(), "-8")
            __check((negative shr 2).toString(), "-16")
            __check((negative shr 1).toString(), "-32")
        }
