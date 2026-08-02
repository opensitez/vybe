// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_copy_of_exact_size
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            val copied = bytes.copyOf()
            __check((copied.joinToString(",")).toString(), "1,2,3")
            __check((copied.contentEquals(bytes)).toString(), "true")
        }
