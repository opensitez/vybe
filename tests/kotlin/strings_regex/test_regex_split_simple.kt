// vybe-test: kotlin/strings_regex/test_regex_split_simple
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\s+")
            val parts = pattern.split("a  b   c")
            __check((parts.size).toString(), "3")
            __check((parts.joinToString("|")).toString(), "a|b|c")
        }
