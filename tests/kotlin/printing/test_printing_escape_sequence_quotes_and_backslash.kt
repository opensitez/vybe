// vybe-test: kotlin/printing/test_printing_escape_sequence_quotes_and_backslash
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("\"q\"").toString(), "\"q\"")
            __check(("a\\b").toString(), "a\\b")
            __check(("\u0041\u0062").toString(), "Ab")
        }
