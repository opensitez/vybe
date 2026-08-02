// vybe-test: kotlin/kotlin_bytes_encoding/test_string_to_byte_array_iso_8859_1
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "Cafe"
            val bytes = source.toByteArray(Charsets.ISO_8859_1)
            val value = String(bytes, Charsets.ISO_8859_1)
            __check((bytes.size).toString(), "4")
            __check((value).toString(), "Cafe")
            __check((bytes.joinToString(",")).toString(), "67,97,102,101")
        }
