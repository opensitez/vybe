// vybe-test: kotlin/bitwise_operations/test_masking_chain_preserves_expected_bits
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0b10101111
            val lowNibble = value and 0x0F
            val upperNibble = (value and 0xF0) ushr 4
            __check((lowNibble).toString(), "15")
            __check((upperNibble).toString(), "10")
            __check((((upperNibble shl 4) or lowNibble)).toString(), "175")
        }
