// vybe-test: kotlin/strings/test_split_with_limit_and_regex_whitespace
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parts = "a b  c d".trim().split("\\s+".toRegex(), 3)
            __check((parts.size).toString(), "3")
            __check((parts[0]).toString(), "a")
            __check((parts[1]).toString(), "b")
            __check((parts[2]).toString(), "c d")
        }
