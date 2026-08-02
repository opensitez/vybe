// vybe-test: kotlin/conversions/test_double_negative_truncation_edges
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((-9.0).toInt()).toString(), "-9")
            __check(((-9.1).toInt()).toString(), "-9")
            __check(((-9.9).toInt()).toString(), "-9")
            __check(((-0.9).toInt()).toString(), "0")
        }
