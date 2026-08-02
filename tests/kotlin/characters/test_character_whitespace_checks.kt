// vybe-test: kotlin/characters/test_character_whitespace_checks
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((' '.isWhitespace()).toString(), "true")
            __check(('\t'.isWhitespace()).toString(), "true")
            __check(('\n'.isWhitespace()).toString(), "true")
            __check(('a'.isWhitespace()).toString(), "false")
        }
