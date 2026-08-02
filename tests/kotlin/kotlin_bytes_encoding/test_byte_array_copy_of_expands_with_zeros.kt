// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_copy_of_expands_with_zeros
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            val expanded = bytes.copyOf(5)
            __check((expanded.joinToString(",")).toString(), "1,2,3,0,0")
            __check((expanded.size).toString(), "5")
        }
