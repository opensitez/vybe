// vybe-test: kotlin/conversions/test_string_to_int_handles_sign_prefixes
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("+12".toInt()).toString(), "12")
            __check(("-12".toInt()).toString(), "-12")
            __check(("0".toInt()).toString(), "0")
        }
