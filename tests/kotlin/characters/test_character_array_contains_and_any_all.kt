// vybe-test: kotlin/characters/test_character_array_contains_and_any_all
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val chars = charArrayOf('k', 'o', 't', 'l', 'i', 'n')
            __check((chars.contains('t')).toString(), "true")
            __check((chars.contains('a')).toString(), "false")
            __check((chars.any { it.isVowel() }).toString(), "true")
            __check((chars.all { it.isLetter() }).toString(), "true")
        }
