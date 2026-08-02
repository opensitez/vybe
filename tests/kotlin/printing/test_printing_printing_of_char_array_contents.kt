// vybe-test: kotlin/printing/test_printing_printing_of_char_array_contents
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val chars = charArrayOf('a', 'b', 'c')
            __check((chars.joinToString("")).toString(), "abc")
        }
