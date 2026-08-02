// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_plus_with_empty_array
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 2, 3) + byteArrayOf()
            __check((values.joinToString(",")).toString(), "1,2,3")
            val empty = byteArrayOf() + byteArrayOf(9, 10)
            __check((empty.joinToString(",")).toString(), "9,10")
        }
