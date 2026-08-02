// vybe-test: kotlin/strings_regex/test_regex_split_limit_one_behavior
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex(",")
            val parts = pattern.split("a,b,c", 1)
            __check((parts.size).toString(), "1")
            __check((parts[0]).toString(), "a,b,c")
        }
