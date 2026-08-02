// vybe-test: kotlin/bitwise_operations/test_bitwise_parity_test_with_masking
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val numbers = listOf(10, 11, 12, 13, 14, 15)
            val parity = numbers.map { it and 1 }
            val even = numbers.filter { it and 1 == 0 }
            __check((parity.joinToString(",")).toString(), "0,1,0,1,0,1")
            __check((even.joinToString(",")).toString(), "10,12,14")
        }
