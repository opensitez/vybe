// vybe-test: kotlin/characters/test_character_mutable_string_append
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = ""
            val a = 'A'
            value += a
            value += ':'
            value += 'B'
            __check((value).toString(), "A:B")
            __check((value.length).toString(), "3")
        }
