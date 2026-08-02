// vybe-test: kotlin/conversions/test_double_parse_infinite_and_nan_keywords
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("Infinity".toDouble()).toString(), "Infinity")
            __check(("-Infinity".toDouble()).toString(), "-Infinity")
            __check(("NaN".toDouble().isNaN()).toString(), "true")
        }
