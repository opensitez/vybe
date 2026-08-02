// vybe-test: kotlin/characters/test_character_count_in_string_parts
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "a1b2c3"
            __check((text.count { it.isDigit() }).toString(), "3")
            __check((text.count { it.isLetter() }).toString(), "3")
            __check((text.count { it.isWhitespace() }).toString(), "0")
        }
