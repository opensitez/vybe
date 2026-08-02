// vybe-test: kotlin/strings_regex/test_regex_find_first_match
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\d+")
            val result = pattern.find("id-42-code")
            __check((result?.value ?: "none").toString(), "42")
            __check((result?.range?.first ?: -1).toString(), "3")
            __check((result?.range?.last ?: -1).toString(), "4")
        }
