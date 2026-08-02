// vybe-test: kotlin/conversions/test_string_to_long_roundtrip
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1234567890123L
            val text = value.toString()
            val parsed = text.toLong()
            __check((text).toString(), "1234567890123")
            __check((parsed == value).toString(), "true")
        }
