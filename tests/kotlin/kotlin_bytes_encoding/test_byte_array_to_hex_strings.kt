// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_to_hex_strings
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(0, 10, 15, 127, -1)
            val hex = values.joinToString(",") {
                val u = it.toInt() and 0xFF
                u.toString(16).padStart(2, '0')
            }
            __check((hex).toString(), "00,0a,0f,7f,ff")
        }
