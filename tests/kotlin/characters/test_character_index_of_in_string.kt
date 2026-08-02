// vybe-test: kotlin/characters/test_character_index_of_in_string
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "banana"
            __check((word.indexOf('a')).toString(), "1")
            __check((word.indexOf('a', 2)).toString(), "3")
            __check((word.lastIndexOf('a')).toString(), "5")
            __check((word.count { it == 'a' }).toString(), "3")
        }
