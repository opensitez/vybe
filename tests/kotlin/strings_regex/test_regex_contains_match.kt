// vybe-test: kotlin/strings_regex/test_regex_contains_match
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("cat|dog")
            __check((pattern.containsMatchIn("the catalog")).toString(), "true")
            __check((pattern.containsMatchIn("fish")).toString(), "false")
        }
