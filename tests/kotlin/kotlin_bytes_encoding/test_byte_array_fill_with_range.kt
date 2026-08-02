// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_fill_with_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = ByteArray(8)
            bytes.fill(1, fromIndex = 2, toIndex = 6)
            __check((bytes.joinToString(",")).toString(), "0,0,1,1,1,1,0,0")
        }
