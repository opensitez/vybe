// vybe-test: kotlin/characters/test_title_case_character_conversion
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('a'.titlecaseChar()).toString(), "A")
            __check(('b'.titlecaseChar()).toString(), "B")
            __check(('1'.titlecaseChar()).toString(), "1")
        }
