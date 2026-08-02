// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_as_sequence_map_to_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(2, 3, 4)
            val total = values.asSequence().map { it.toInt() }.sum()
            __check((total).toString(), "9")
            __check((values.asSequence().map { it + 1 }.toList().joinToString(",")).toString(), "3,4,5")
        }
