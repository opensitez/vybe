// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_factory_size_and_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(7, 8, 9, 10)
            __check((bytes.size).toString(), "4")
            __check((bytes.joinToString(",")).toString(), "7,8,9,10")
        }
