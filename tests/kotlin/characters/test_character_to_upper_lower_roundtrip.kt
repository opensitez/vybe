// vybe-test: kotlin/characters/test_character_to_upper_lower_roundtrip
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('x'.uppercaseChar().lowercaseChar()).toString(), "x")
            __check(('X'.lowercaseChar().uppercaseChar()).toString(), "X")
            __check(('ß'.lowercaseChar()).toString(), "ß")
            __check(('ß'.uppercaseChar()).toString(), "SS")
        }
