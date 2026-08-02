// vybe-test: kotlin/characters/test_character_building_and_mutable_lists
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val chars = mutableListOf<Char>()
            chars.add('k')
            chars.add('o')
            chars.add('t')
            chars[1] = 'O'
            __check((chars.joinToString("")).toString(), "kOt")
            __check((chars.size).toString(), "3")
            __check((chars[2]).toString(), "t")
        }
