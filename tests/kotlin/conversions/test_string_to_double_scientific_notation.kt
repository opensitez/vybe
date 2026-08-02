// vybe-test: kotlin/conversions/test_string_to_double_scientific_notation
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("1.5e2".toDouble()).toString(), "150")
            __check(("2.5E-1".toDouble()).toString(), "0.25")
        }
