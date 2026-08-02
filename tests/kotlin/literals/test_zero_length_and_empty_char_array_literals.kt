// vybe-test: kotlin/literals/test_zero_length_and_empty_char_array_literals
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = ""
            val letters = charArrayOf()
            __check((empty.isEmpty()).toString(), "true")
            __check((letters.isEmpty()).toString(), "true")
            __check((empty == "").toString(), "true")
            __check((letters.size).toString(), "0")
        }
