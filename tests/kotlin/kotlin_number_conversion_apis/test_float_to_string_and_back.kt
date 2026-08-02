// vybe-test: kotlin/kotlin_number_conversion_apis/test_float_to_string_and_back
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 12.5
            val text = value.toString()
            __check((text).toString(), "12.5")
            __check((text.toDouble()).toString(), "12.5")
            __check((2.0.toString()).toString(), "2.0")
            __check((2.0.toString().toIntOrNull()).toString(), "null")
        }
