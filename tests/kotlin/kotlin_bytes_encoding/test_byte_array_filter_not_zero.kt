// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_filter_not_zero
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(0, 1, 0, 2, 3)
            val filtered = values.filterNot { it == 0 }
            __check((filtered.joinToString(",")).toString(), "1,2,3")
        }
