// vybe-test: kotlin/kotlin_char_predicates/test_char_classification_methods
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_predicates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('A'.isLetter()).toString(), "true")
            __check(('3'.isDigit()).toString(), "true")
            __check((' '.isWhitespace()).toString(), "true")
        }
