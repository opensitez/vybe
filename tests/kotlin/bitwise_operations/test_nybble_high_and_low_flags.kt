// vybe-test: kotlin/bitwise_operations/test_nybble_high_and_low_flags
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0b11010101
            __check((value and 0b1111).toString(), "5")
            __check((value shr 4).toString(), "13")
            __check((value and 0b11110000).toString(), "208")
            __check(((value and 0b11110000) shr 4).toString(), "13")
        }
