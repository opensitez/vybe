// vybe-test: kotlin/characters/test_character_digit_validation
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val chars = listOf('1', '9', 'a', ' ')
            __check((chars.all { it.isDigit() }).toString(), "false")
            __check((chars.count { it.isDigit() }).toString(), "2")
            __check((chars[0].digitToInt()).toString(), "1")
            __check((chars[1].digitToInt()).toString(), "9")
            __check((chars[2].digitToIntOrNull() == null).toString(), "true")
        }
