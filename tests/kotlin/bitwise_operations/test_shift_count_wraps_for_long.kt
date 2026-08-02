// vybe-test: kotlin/bitwise_operations/test_shift_count_wraps_for_long
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1L shl 63).toString(), "-9223372036854775808")
            __check((1L shl 64).toString(), "1")
            __check((1L shl 65).toString(), "2")
            __check((1L shl -1).toString(), "-9223372036854775808")
        }
