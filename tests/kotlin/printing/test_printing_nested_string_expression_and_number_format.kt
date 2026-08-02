// vybe-test: kotlin/printing/test_printing_nested_string_expression_and_number_format
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val size = 3
            val label = "count=$size"
            __check((label).toString(), "count=3")
            __check(("${label.length}").toString(), "7")
        }
