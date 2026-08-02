// vybe-test: kotlin/printing/test_printing_character_singleton
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('A').toString(), "A")
            __check(('9').toString(), "9")
        }
