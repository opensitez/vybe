// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_empty_copy_of_range_behavior
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            val empty = bytes.copyOfRange(1, 1)
            __check((empty.size).toString(), "0")
            __check((empty.isEmpty()).toString(), "true")
        }
