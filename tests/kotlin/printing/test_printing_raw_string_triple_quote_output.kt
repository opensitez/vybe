// vybe-test: kotlin/printing/test_printing_raw_string_triple_quote_output
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val raw = """line1
line2
line3"""
            __check((raw).toString(), "line1\nline2\nline3")
        }
