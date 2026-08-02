// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_copy_of_truncates
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(1, 2, 3, 4, 5)
            val truncated = bytes.copyOf(3)
            __check((truncated.joinToString(",")).toString(), "1,2,3")
            __check((truncated.size).toString(), "3")
        }
