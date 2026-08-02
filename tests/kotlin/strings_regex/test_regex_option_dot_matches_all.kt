// vybe-test: kotlin/strings_regex/test_regex_option_dot_matches_all
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val dot = Regex("a.*c")
            val dotAny = Regex("a.*c", RegexOption.DOT_MATCHES_ALL)
            val text = "a\nc"
            __check((dot.matches(text)).toString(), "false")
            __check((dotAny.matches(text)).toString(), "true")
        }
