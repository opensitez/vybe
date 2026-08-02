// vybe-test: kotlin/printing/test_printing_prints_multi_line_raw_template
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val label = "items"
            __check(("""$label:
  - one
  - two""").toString(), "items:\n  - one\n  - two")
        }
