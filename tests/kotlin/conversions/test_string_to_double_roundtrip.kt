// vybe-test: kotlin/conversions/test_string_to_double_roundtrip
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "3.75"
            val value = source.toDouble()
            __check((value).toString(), "3.75")
            __check((value.toString()).toString(), "3.75")
        }
