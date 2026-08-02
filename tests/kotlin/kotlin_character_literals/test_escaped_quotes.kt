// vybe-test: kotlin/kotlin_character_literals/test_escaped_quotes
// origin: languages/kotlin/tests/kotlin/test_kotlin_character_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val quote = '\''
            val backslash = '\\'
            __check((quote).toString(), "'")
            __check((backslash).toString(), "\\")
        }
