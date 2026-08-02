// vybe-test: kotlin/strings_regex/test_string_to_regex_matches
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val number = "\\d{2,3}".toRegex()
            __check((number.matches("42")).toString(), "true")
            __check((number.matches("4")).toString(), "false")
        }
