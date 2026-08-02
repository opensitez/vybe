// vybe-test: kotlin/strings_regex/test_regex_match_and_non_match
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val number = Regex("\\d+")
            __check((number.matches("12345")).toString(), "true")
            __check((number.matches("12a45")).toString(), "false")
        }
