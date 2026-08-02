// vybe-test: kotlin/strings_regex/test_regex_literal_option_treats_regex_meta_as_text
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("a+b|c*", RegexOption.LITERAL)
            __check((pattern.matches("a+b|c*")).toString(), "true")
            __check((pattern.matches("aaab")).toString(), "false")
            __check((pattern.containsMatchIn("xxa+b|c*yy")).toString(), "true")
        }
