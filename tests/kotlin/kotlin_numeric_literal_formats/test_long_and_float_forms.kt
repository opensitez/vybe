// vybe-test: kotlin/kotlin_numeric_literal_formats/test_long_and_float_forms
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_literal_formats.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val big = 1_000_000L
            val small = 1.5
            __check((big).toString(), "1000000")
            __check((small.toString()).toString(), "1.5")
        }
