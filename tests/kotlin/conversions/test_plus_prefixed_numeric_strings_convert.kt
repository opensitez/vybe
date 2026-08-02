// vybe-test: kotlin/conversions/test_plus_prefixed_numeric_strings_convert
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("+42".toInt()).toString(), "42")
            __check(("+3.25".toDouble()).toString(), "3.25")
            __check(("0007".toInt()).toString(), "7")
        }
