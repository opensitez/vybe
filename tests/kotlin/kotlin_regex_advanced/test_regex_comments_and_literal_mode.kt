// vybe-test: kotlin/kotlin_regex_advanced/test_regex_comments_and_literal_mode
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val token = "a#b".toRegex(setOf(RegexOption.COMMENTS))
            val literal = Regex("a#b", RegexOption.LITERAL)
            __check((token.matches("a#b")).toString(), "false")
            __check((literal.matches("a#b")).toString(), "true")
            __check((token.matches("a b")).toString(), "false")
        }
