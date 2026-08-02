// vybe-test: kotlin/characters/test_character_unicode_literal
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val copyright = '\u00A9'
            val omega = '\u03A9'
            __check((copyright).toString(), "©")
            __check((omega).toString(), "Ω")
        }
