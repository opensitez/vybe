// vybe-test: kotlin/conversions/test_double_to_int_truncation
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((9.9.toInt()).toString(), "9")
            __check(((-9.9).toInt()).toString(), "-9")
        }
