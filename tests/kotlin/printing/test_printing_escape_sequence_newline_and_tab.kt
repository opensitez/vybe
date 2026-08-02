// vybe-test: kotlin/printing/test_printing_escape_sequence_newline_and_tab
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("a\n b").toString(), "a\n b")
            __check(("x\t y").toString(), "x\t y")
        }
