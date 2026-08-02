// vybe-test: kotlin/conversions/test_string_to_long_and_negative_roundtrip
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = -987654321L
            val parsed = value.toString().toLong()
            __check((parsed).toString(), "-987654321")
            __check((parsed == value).toString(), "true")
        }
