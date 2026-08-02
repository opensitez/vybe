// vybe-test: kotlin/bitwise_operations/test_bitwise_clear_all_low_bits
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0xFFFF
            val cleared = value and 0xFFF0
            val low = value and 0x000F
            __check((cleared).toString(), "65520")
            __check((low).toString(), "15")
            __check(((cleared and low)).toString(), "0")
        }
