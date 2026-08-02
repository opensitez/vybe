// vybe-test: kotlin/kotlin_bytes_encoding/test_string_to_byte_array_round_trip_utf8_unicode
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "hé"
            val bytes = source.toByteArray(Charsets.UTF_8)
            val value = String(bytes, Charsets.UTF_8)
            __check((bytes.size).toString(), "3")
            __check((value).toString(), "hé")
            __check((bytes.first()).toString(), "-61")
        }
