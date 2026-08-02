// vybe-test: kotlin/kotlin_string_number_parsing/test_parse_long_with_sign
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("-12".toLong()).toString(), "-12")
            __check(("+12".toLong()).toString(), "12")
            __check(("12L".toLongOrNull()).toString(), "null")
        }
