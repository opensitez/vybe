// vybe-test: kotlin/bitwise_operations/test_bitwise_filters_using_shifted_masks
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(0, 1, 2, 3, 4, 5, 6, 7, 8, 15, 16, 31)
            val maskedTwoBits = values.map { it and 0b11 }
            val flags = values.filter { (it and 0b1000) == 0b1000 }
            __check((maskedTwoBits.joinToString(",")).toString(), "0,1,2,3,0,1,2,3,0,3,0,3")
            __check((flags.joinToString(",")).toString(), "8,15")
        }
