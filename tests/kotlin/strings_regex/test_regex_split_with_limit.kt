// vybe-test: kotlin/strings_regex/test_regex_split_with_limit
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex(",")
            val parts = pattern.split("a,b,c,d", limit = 2)
            __check((parts.size).toString(), "2")
            __check((parts[0]).toString(), "a")
            __check((parts[1]).toString(), "b,c,d")
        }
