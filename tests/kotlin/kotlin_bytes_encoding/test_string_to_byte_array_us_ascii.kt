// vybe-test: kotlin/kotlin_bytes_encoding/test_string_to_byte_array_us_ascii
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "abc123"
            val bytes = source.toByteArray(Charsets.US_ASCII)
            val value = String(bytes, Charsets.US_ASCII)
            __check((bytes.size).toString(), "6")
            __check((value).toString(), "abc123")
            __check((bytes.sum()).toString(), "594")
        }
