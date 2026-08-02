// vybe-test: kotlin/strings_regex/test_regex_match_entire_boundary
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("[a-z]+")
            __check((pattern.matchEntire("abc") != null).toString(), "true")
            __check((pattern.matchEntire("abc123") != null).toString(), "false")
            __check((pattern.matches("abc")).toString(), "true")
        }
