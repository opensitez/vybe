// vybe-test: kotlin/characters/test_character_filter_with_unicode_category
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = "a1b2c3"
            val letters = input.filter { it.isLetter() }
            val digits = input.filter { it.isDigit() }
            val mapped = input.map { if (it.isDigit()) '*' else it }
            __check((letters).toString(), "abc")
            __check((digits).toString(), "123")
            __check((mapped.joinToString("")).toString(), "a*b*c*")
        }
