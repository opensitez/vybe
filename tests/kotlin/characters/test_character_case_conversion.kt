// vybe-test: kotlin/characters/test_character_case_conversion
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('k'.uppercaseChar()).toString(), "K")
            __check(('M'.lowercaseChar()).toString(), "m")
            __check(('ß'.uppercaseChar()).toString(), "SS")
            __check(('ß'.lowercaseChar()).toString(), "ß")
        }
