// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_copy_of_range_middle_slice
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(10, 20, 30, 40, 50)
            val middle = bytes.copyOfRange(1, 4)
            __check((middle.joinToString(",")).toString(), "20,30,40")
            __check((middle.size).toString(), "3")
        }
