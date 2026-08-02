// vybe-test: kotlin/kotlin_number_parsing/test_parse_numbers_in_expected_formats
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("12".toInt()).toString(), "12")
            __check(("3.5".toDouble()).toString(), "3.5")
            __check(("ff".toInt(16)).toString(), "255")
        }
