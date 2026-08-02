// vybe-test: kotlin/kotlin_regex_advanced/test_regex_find_and_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val regex = Regex("b")
            val match = regex.find("abca", 0)
            __check((match?.value).toString(), "b")
            __check((match?.range?.first).toString(), "1")
            __check((match?.range?.last).toString(), "1")
        }
