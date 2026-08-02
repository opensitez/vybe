// vybe-test: kotlin/bitwise_operations/test_long_bitwise_and
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Long = 0xFFL
            val mask: Long = 0x0F
            __check((value and mask).toString(), "15")
            __check(((value xor mask)).toString(), "240")
            __check((value or mask).toString(), "255")
        }
