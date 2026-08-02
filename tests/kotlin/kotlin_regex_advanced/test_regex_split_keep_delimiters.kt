// vybe-test: kotlin/kotlin_regex_advanced/test_regex_split_keep_delimiters
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val regex = Regex("[;,]")
            val parts = regex.split("a,b;c,d")
            __check((parts.joinToString("|")).toString(), "a|b|c|d")
            __check((parts.size).toString(), "4")
        }
