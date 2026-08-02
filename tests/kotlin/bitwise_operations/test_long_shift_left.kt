// vybe-test: kotlin/bitwise_operations/test_long_shift_left
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Long = 1L
            __check((value shl 8).toString(), "256")
            __check(((1L shl 32).toString()).toString(), "4294967296")
            __check(((3L shl 5)).toString(), "96")
        }
