// vybe-test: kotlin/bitwise_operations/test_bitwise_toggle_odd_bits
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5, 6, 7)
            val toggled = values.map { it xor 1 }
            __check((toggled.joinToString(",")).toString(), "0,3,2,5,4,7,6")
            val restored = toggled.map { it xor 1 }
            __check((restored.joinToString(",")).toString(), "1,2,3,4,5,6,7")
        }
