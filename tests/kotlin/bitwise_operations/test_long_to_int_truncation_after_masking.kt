// vybe-test: kotlin/bitwise_operations/test_long_to_int_truncation_after_masking
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val wide: Long = 0x1_0000_0000L
            val narrowed = (wide and 0xFFFF_FFFF).toInt()
            __check((wide.toString()).toString(), "4294967296")
            __check((narrowed).toString(), "0")
            __check((wide.toInt()).toString(), "0")
        }
