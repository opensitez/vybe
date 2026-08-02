// vybe-test: kotlin/kotlin_string_number_parsing/test_parse_double_and_float
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("12.75".toDouble()).toString(), "12.75")
            __check(("-0.5".toFloat()).toString(), "-0.5")
            __check(("nan".toDoubleOrNull()?.isNaN() ?: false).toString(), "true")
            __check(("bad".toDoubleOrNull()).toString(), "null")
        }
