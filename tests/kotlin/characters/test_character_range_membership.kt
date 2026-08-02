// vybe-test: kotlin/characters/test_character_range_membership
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('b' in 'a'..'f').toString(), "true")
            __check(('z' in 'a'..'f').toString(), "false")
            __check(('5' in '0'..'9').toString(), "true")
            __check(('g' in 'a'..'f').toString(), "false")
        }
