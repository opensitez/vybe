// vybe-test: kotlin/characters/test_character_comparison_to_string_length
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = "ab"
            val right = "aB"
            __check((left[0] < right[1]).toString(), "false")
            __check((left[1] == 'b').toString(), "true")
            __check((right.compareTo("aa")).toString(), "1")
            __check((left.compareTo(right)).toString(), "-1")
        }
