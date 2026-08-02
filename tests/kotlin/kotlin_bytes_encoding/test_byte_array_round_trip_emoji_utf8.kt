// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_round_trip_emoji_utf8
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "🙂"
            val bytes = source.toByteArray(Charsets.UTF_8)
            val value = String(bytes, Charsets.UTF_8)
            __check((bytes.size).toString(), "4")
            __check((value).toString(), "🙂")
            __check((value == source).toString(), "true")
        }
