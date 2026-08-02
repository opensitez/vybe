// vybe-test: kotlin/characters/test_character_slice_from_indexes
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "abcdef"
            __check((word[0]).toString(), "a")
            __check((word[3]).toString(), "d")
            __check((word[5]).toString(), "f")
            __check((word[word.lastIndex]).toString(), "f")
        }
