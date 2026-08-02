// vybe-test: kotlin/strings_regex/test_regex_grouping_and_replacement_with_backreference
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("(\\w+)\\s+\\1")
            __check((pattern.matches("go go")).toString(), "true")
            __check((pattern.matches("go now")).toString(), "false")
            val text = pattern.replace("yo yo test") { match ->
                "[${match.groupValues[1]}]"
            }
            __check((text).toString(), "[yo] test")
        }
