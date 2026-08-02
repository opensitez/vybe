// vybe-test: kotlin/kotlin_regex_advanced/test_regex_dot_matches_all
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "a\nb"
            val normal = Regex("a.b")
            val dotAll = Regex("a.b", RegexOption.DOT_MATCHES_ALL)
            __check((normal.containsMatchIn(text)).toString(), "false")
            __check((dotAll.containsMatchIn(text)).toString(), "true")
        }
