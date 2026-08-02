// vybe-test: kotlin/characters/test_single_character_output
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('a').toString(), "a")
            __check(('Z').toString(), "Z")
        }
