// vybe-test: kotlin/bitwise_operations/test_long_unsigned_shift_right
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val negative: Long = -1L
            val signed: Long = -16L
            __check((negative ushr 1).toString(), "9223372036854775807")
            __check((signed ushr 2).toString(), "2305843009213693950")
            __check((15L ushr 1).toString(), "7")
            __check((1L ushr 1).toString(), "0")
        }
