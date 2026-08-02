// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_count_and_first_last
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(8, 0, 8, 4)
            __check((values.count { it == 8 }).toString(), "2")
            __check((values.first()).toString(), "8")
            __check((values.last()).toString(), "4")
        }
