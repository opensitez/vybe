// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_sum_and_reduce
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 2, 3)
            __check((values.sum()).toString(), "6")
            __check((values.reduce { acc, value -> acc + value }).toString(), "6")
            __check((values.fold(0) { acc, value -> acc + value }).toString(), "6")
        }
