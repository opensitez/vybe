// vybe-test: kotlin/strings_regex/test_regex_unicode_class_and_case_option_interaction
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val letters = Regex("straße", RegexOption.IGNORE_CASE)
            __check((letters.matches("STRASSE")).toString(), "true")
            __check((letters.matchesAt("XXstraßeYY", 2)).toString(), "true")
        }
