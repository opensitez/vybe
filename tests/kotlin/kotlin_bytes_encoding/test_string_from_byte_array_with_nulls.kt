// vybe-test: kotlin/kotlin_bytes_encoding/test_string_from_byte_array_with_nulls
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(78, 117, 108, 108, 0, 65)
            val value = String(bytes)
            __check((value.length).toString(), "6")
            __check((value[0]).toString(), "N")
            __check((value[4].code).toString(), "0")
        }
