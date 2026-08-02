// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_join_to_string_formatting
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(3, 1, 4)
            __check((values.joinToString("|")).toString(), "3|1|4")
            __check((values.joinToString(",", prefix = "[", postfix = "]")).toString(), "[3,1,4]")
            __check((values.joinToString(";", transform = { it.toString() })).toString(), "3;1;4")
        }
