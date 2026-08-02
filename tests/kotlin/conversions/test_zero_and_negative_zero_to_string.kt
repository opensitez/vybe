// vybe-test: kotlin/conversions/test_zero_and_negative_zero_to_string
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0.toString()).toString(), "0")
            __check(((-0).toDouble().toInt()).toString(), "0")
            __check(((-0.0).toString()).toString(), "0")
        }
