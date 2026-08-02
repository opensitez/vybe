// vybe-test: kotlin/kotlin_char_apis/test_char_identifier_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('a'.isLetter()).toString(), "true")
            __check((''.isLetter()).toString(), "false")
            __check(('7'.isDigit()).toString(), "true")
            __check(('_'.isIdentifierStart()).toString(), "true")
            __check(('x'.isIdentifierPart()).toString(), "true")
            __check((' '.isIdentifierPart()).toString(), "false")
        }
