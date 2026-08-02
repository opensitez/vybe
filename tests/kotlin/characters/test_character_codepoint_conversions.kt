// vybe-test: kotlin/characters/test_character_codepoint_conversions
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 'a'
            val b = 'b'
            __check((a.code + 1).toString(), "98")
            __check(((b.code - a.code)).toString(), "1")
            __check((97.toChar()).toString(), "a")
            __check((('A'.code - 'a'.code)).toString(), "-32")
        }
