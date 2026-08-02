// vybe-test: kotlin/strings_regex/test_regex_matches_at_offset
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\d+")
            __check((pattern.matchesAt("abc123", 3)).toString(), "true")
            __check((pattern.matchesAt("abc123", 0)).toString(), "false")
        }
