// vybe-test: kotlin/kotlin_regex_advanced/test_regex_ignore_case_option
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val re = "ab+c".toRegex(RegexOption.IGNORE_CASE)
            __check((re.matches("ABBC")).toString(), "true")
            __check((re.matches("abc")).toString(), "true")
            __check((re.matches("ABC")).toString(), "true")
        }
