// vybe-test: kotlin/kotlin_character_literals/test_digit_and_letter_chars
// origin: languages/kotlin/tests/kotlin/test_kotlin_character_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c1 = '1'
            val c2 = 'A'
            __check((c1.toString() + c2.toString()).toString(), "1A")
        }
