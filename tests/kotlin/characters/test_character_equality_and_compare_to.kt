// vybe-test: kotlin/characters/test_character_equality_and_compare_to
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('k' == 'k').toString(), "true")
            __check(('k' != 'K').toString(), "true")
            __check(('a' < 'c').toString(), "true")
            __check(('z'.compareTo('a')).toString(), "25")
            __check(('a'.compareTo('z')).toString(), "-25")
        }
