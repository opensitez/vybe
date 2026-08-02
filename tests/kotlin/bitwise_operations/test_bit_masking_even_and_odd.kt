// vybe-test: kotlin/bitwise_operations/test_bit_masking_even_and_odd
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mask = 1
            val values = listOf(1, 2, 3, 4, 5, 6, 7, 8)
            val onlyEven = values.filter { it and 1 == 0 }
            val onlyOdd = values.filter { it and 1 == 1 }
            __check((onlyEven.joinToString(",")).toString(), "2,4,6,8")
            __check((onlyOdd.joinToString(",")).toString(), "1,3,5,7")
        }
