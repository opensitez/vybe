// vybe-test: kotlin/string_search_apis/test_index_of_any_character_set
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "apple,banana,carrot"
            __check((text.indexOfAny(charArrayOf(',', 'a'), 0)).toString(), "5")
            __check((text.indexOfAny(charArrayOf('-', '/'))).toString(), "-1")
        }
