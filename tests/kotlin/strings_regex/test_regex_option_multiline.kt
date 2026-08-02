// vybe-test: kotlin/strings_regex/test_regex_option_multiline
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val anchored = Regex("^kotlin$", RegexOption.MULTILINE)
            __check((anchored.containsMatchIn("java\nkotlin\nrust")).toString(), "true")
            __check((anchored.containsMatchIn("kotlin ")).toString(), "false")
        }
