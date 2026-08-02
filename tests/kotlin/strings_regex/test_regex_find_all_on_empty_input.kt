// vybe-test: kotlin/strings_regex/test_regex_find_all_on_empty_input
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\d+")
            val values = pattern.findAll("").toList()
            __check((values.size).toString(), "0")
        }
