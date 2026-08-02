// vybe-test: kotlin/strings_regex/test_regex_find_on_starting_index
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\d+")
            val first = pattern.find("a1b22", 2)
            __check((first?.value ?: "none").toString(), "22")
        }
