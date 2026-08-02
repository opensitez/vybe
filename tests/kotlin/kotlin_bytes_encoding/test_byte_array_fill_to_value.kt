// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_fill_to_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = ByteArray(5)
            bytes.fill(7)
            __check((bytes.joinToString(",")).toString(), "7,7,7,7,7")
        }
