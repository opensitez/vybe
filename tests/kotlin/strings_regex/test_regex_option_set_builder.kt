// vybe-test: kotlin/strings_regex/test_regex_option_set_builder
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("k.t", setOf(RegexOption.IGNORE_CASE, RegexOption.DOT_MATCHES_ALL))
            __check((pattern.matches("K\nt")).toString(), "true")
        }
