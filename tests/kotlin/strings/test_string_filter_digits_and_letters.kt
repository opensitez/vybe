// vybe-test: kotlin/strings/test_string_filter_digits_and_letters
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "a1b2c3"
            val digits = value.filter { it.isDigit() }
            val letters = value.filter { it.isLetter() }
            __check((digits).toString(), "123")
            __check((letters).toString(), "abc")
        }
