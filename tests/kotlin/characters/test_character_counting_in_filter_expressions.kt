// vybe-test: kotlin/characters/test_character_counting_in_filter_expressions
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "A1 b2C-3"
            __check((value.count { it.isUpperCase() }).toString(), "2")
            __check((value.count { it.isLowerCase() }).toString(), "1")
            __check((value.count { it.isDigit() }).toString(), "3")
            __check((value.count { it.isWhitespace() }).toString(), "1")
            __check((value.count { it.isLetter() }).toString(), "3")
            __check((value.count { it.isLetterOrDigit() }).toString(), "6")
        }
