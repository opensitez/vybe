// vybe-test: kotlin/characters/test_character_join_to_string_with_transform
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = charArrayOf('a', 'b', 'c')
            val transformed = values.joinToString(",") { it.uppercaseChar().toString() }
            val mapped = values.map { it.uppercaseChar() }.joinToString("")
            __check((transformed).toString(), "A,B,C")
            __check((mapped).toString(), "ABC")
        }
