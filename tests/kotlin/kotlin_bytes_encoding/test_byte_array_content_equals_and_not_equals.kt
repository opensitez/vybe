// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_content_equals_and_not_equals
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = byteArrayOf(1, 2, 3)
            val b = byteArrayOf(1, 2, 3)
            val c = byteArrayOf(3, 2, 1)
            __check((a.contentEquals(b)).toString(), "true")
            __check((a.contentEquals(c)).toString(), "false")
            __check((a.contentEquals(a)).toString(), "true")
        }
