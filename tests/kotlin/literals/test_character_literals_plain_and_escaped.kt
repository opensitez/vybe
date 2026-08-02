// vybe-test: kotlin/literals/test_character_literals_plain_and_escaped
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val letter = 'K'
            val quote = '\''
            val backslash = '\\'
            __check((letter).toString(), "K")
            __check((quote).toString(), "'")
            __check((backslash).toString(), "\\")
        }
