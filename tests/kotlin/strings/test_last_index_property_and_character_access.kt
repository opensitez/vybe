// vybe-test: kotlin/strings/test_last_index_property_and_character_access
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "rust"
            __check((word.lastIndex).toString(), "3")
            __check((word[0]).toString(), "r")
            __check((word[word.lastIndex]).toString(), "t")
        }
