// vybe-test: kotlin/characters/test_character_to_string_and_joining
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val chars = charArrayOf('a', 'b', 'c')
            __check((chars.joinToString(",")).toString(), "a,b,c")
            __check((chars[0].toString()).toString(), "a")
        }
