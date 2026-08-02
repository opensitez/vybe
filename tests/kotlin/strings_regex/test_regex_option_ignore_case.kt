// vybe-test: kotlin/strings_regex/test_regex_option_ignore_case
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val lower = Regex("cat", RegexOption.IGNORE_CASE)
            __check((lower.matches("CAT")).toString(), "true")
            __check((lower.matches("dog")).toString(), "false")
        }
