// vybe-test: kotlin/kotlin_regex_advanced/test_regex_match_entire_line_with_options_combo
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("^\n*OK\\?$")
            val withComments = Regex("^\n*OK\\?$")
            __check((pattern.matches("OK?")).toString(), "true")
            __check((withComments.matches("\n\nOK?")).toString(), "false")
            val withOption = Regex("^\n*OK\\?$", RegexOption.MULTILINE)
            __check((withOption.matches("line1\nOK?")).toString(), "false")
        }
