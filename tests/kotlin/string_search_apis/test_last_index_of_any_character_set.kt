// vybe-test: kotlin/string_search_apis/test_last_index_of_any_character_set
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "abXcdY"
            __check((text.lastIndexOfAny(charArrayOf('X', 'Y', 'Z'))).toString(), "5")
            __check((text.lastIndexOfAny(charArrayOf('Q'))).toString(), "-1")
        }
