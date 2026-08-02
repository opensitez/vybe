// vybe-test: kotlin/strings_regex/test_regex_find_all_numbers
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\d+")
            val values = pattern.findAll("a1b22c333")
            __check((values.count()).toString(), "3")
            __check((values.joinToString("|") { it.value }).toString(), "1|22|333")
        }
