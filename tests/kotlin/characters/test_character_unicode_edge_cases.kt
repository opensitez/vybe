// vybe-test: kotlin/characters/test_character_unicode_edge_cases
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val space = '\u0020'
            val tab = '\u0009'
            val euro = '€'
            __check((space.isWhitespace()).toString(), "true")
            __check((tab.isWhitespace()).toString(), "true")
            __check((euro.isLetterOrDigit()).toString(), "false")
            __check((euro.isDefined()).toString(), "true")
            __check((euro.code).toString(), "8364")
        }
