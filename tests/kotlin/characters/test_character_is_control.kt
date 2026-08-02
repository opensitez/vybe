// vybe-test: kotlin/characters/test_character_is_control
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('\u0000'.isISOControl()).toString(), "true")
            __check(('\n'.isISOControl()).toString(), "true")
            __check(('a'.isISOControl()).toString(), "false")
            __check(('\u0009'.isISOControl()).toString(), "true")
        }
