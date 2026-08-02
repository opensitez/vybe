// vybe-test: kotlin/kotlin_number_conversion_apis/test_integer_to_string_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 123
            val s = value.toString()
            __check((s).toString(), "123")
            __check((s.toInt()).toString(), "123")
            __check(((-45).toString()).toString(), "-45")
            __check(((-45).toString().toIntOrNull()).toString(), "-45")
        }
