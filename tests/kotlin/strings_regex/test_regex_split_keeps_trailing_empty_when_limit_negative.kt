// vybe-test: kotlin/strings_regex/test_regex_split_keeps_trailing_empty_when_limit_negative
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex(",")
            val parts = pattern.split("a,b,c,", limit = -1)
            __check((parts.size).toString(), "4")
            __check((parts.joinToString("|")).toString(), "a|b|c|")
        }
