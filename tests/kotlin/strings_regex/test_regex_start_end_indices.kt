// vybe-test: kotlin/strings_regex/test_regex_start_end_indices
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("dog")
            val result = pattern.find("The dog runs")
            __check((result?.range?.start ?: -1).toString(), "4")
            __check((result?.range?.endInclusive ?: -1).toString(), "6")
        }
