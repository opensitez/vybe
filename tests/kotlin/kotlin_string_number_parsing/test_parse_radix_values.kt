// vybe-test: kotlin/kotlin_string_number_parsing/test_parse_radix_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("ff".toInt(16)).toString(), "255")
            __check(("11".toInt(2)).toString(), "3")
            __check(("77".toInt(8)).toString(), "63")
            __check(("123".toIntOrNull(37)).toString(), "null")
        }
