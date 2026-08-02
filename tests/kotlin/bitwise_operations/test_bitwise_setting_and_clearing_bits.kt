// vybe-test: kotlin/bitwise_operations/test_bitwise_setting_and_clearing_bits
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0
            val withBit2 = value or (1 shl 2)
            val withBit3 = withBit2 or (1 shl 3)
            val cleared = withBit3 and (1 shl 2).inv()
            __check((withBit2).toString(), "4")
            __check((withBit3).toString(), "12")
            __check((cleared).toString(), "8")
        }
