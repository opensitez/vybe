// vybe-test: kotlin/printing/test_printing_range_to_string_form
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1..3).toString(), "1..3")
        }
