// vybe-test: kotlin/bitwise_operations/test_bitwise_counting_subset_flags
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(0b1010, 0b1111, 0b1000, 0b0011)
            val countAnyHigh = values.count { it and 0b1000 != 0 }
            val countZeroLow = values.count { it and 1 == 0 }
            val countPairs = values.filter { (it and 0b0110) == 0b0010 }
            __check((countAnyHigh).toString(), "3")
            __check((countZeroLow).toString(), "3")
            __check((countPairs.joinToString(",")).toString(), "10")
        }
