// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_filter_and_all_any
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 2, 3, 4, 5)
            __check((values.filter { it % 2 == 0 }.joinToString(",")).toString(), "2,4")
            __check((values.any { it > 4 }).toString(), "true")
            __check((values.all { it > 0 }).toString(), "true")
            __check((values.none { it < 0 }).toString(), "true")
        }
