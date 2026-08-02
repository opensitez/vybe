// vybe-test: kotlin/kotlin_regex_advanced/test_regex_multiline_anchors
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "x\ny\nz"
            val withAnchors = "^y$".toRegex(RegexOption.MULTILINE)
            val matched = withAnchors.containsMatchIn(text)
            val bad = "^y$".toRegex().containsMatchIn(text)
            __check((matched).toString(), "true")
            __check((bad).toString(), "false")
        }
