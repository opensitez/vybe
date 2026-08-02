// vybe-test: kotlin/characters/test_character_letter_or_digit_predicates
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('5'.isDigit()).toString(), "true")
            __check(('X'.isLetter()).toString(), "true")
            __check(('X'.isLetterOrDigit()).toString(), "true")
            __check(('#'.isLetterOrDigit()).toString(), "false")
        }
