// vybe-test: kotlin/strings_regex/test_regex_find_with_start_index_and_no_match
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\d+")
            __check((pattern.find("abc123", 1)?.value ?: "none").toString(), "123")
            __check((pattern.find("abc123", 4) == null).toString(), "true")
            __check((pattern.find("abc", 3) == null).toString(), "true")
            __check((pattern.find("abc", 5) == null).toString(), "true")
        }
