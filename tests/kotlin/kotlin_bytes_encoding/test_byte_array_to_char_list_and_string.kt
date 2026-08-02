// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_to_char_list_and_string
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = "AZ09".toByteArray()
            val chars = values.map { it.toInt().toChar() }.joinToString(",")
            __check((chars).toString(), "A,Z,0,9")
            val rebuilt = String(byteArrayOf(65, 90, 48, 57))
            __check((rebuilt).toString(), "AZ09")
        }
