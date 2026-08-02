// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_count_even_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 2, 3, 4, 5, 6)
            __check((values.count { it % 2 == 0 }).toString(), "3")
        }
