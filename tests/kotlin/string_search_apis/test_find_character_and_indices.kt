// vybe-test: kotlin/string_search_apis/test_find_character_and_indices
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "kotlin"
            __check((text.indexOfFirst { it in 'a'..'z' }).toString(), "0")
            __check((text.indexOfLast { it == 't' }).toString(), "2")
            __check((text.find { it == 'i' }).toString(), "k")
            __check((text.findLast { it == 'o' } ?: "none").toString(), "o")
        }
