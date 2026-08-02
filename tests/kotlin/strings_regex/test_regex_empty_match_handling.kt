// vybe-test: kotlin/strings_regex/test_regex_empty_match_handling
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("a*")
            val result = pattern.matchEntire("")
            __check((result != null).toString(), "true")
            __check((result?.value ?: "none").toString(), "")
        }
