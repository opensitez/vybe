// vybe-test: kotlin/bitwise_operations/test_bitwise_roundtrip_with_shift_and_or
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = 0b101010
            val shifted = original shl 2
            val restored = (shifted shr 2) or (original and 0)
            __check((shifted).toString(), "168")
            __check((restored).toString(), "42")
        }
