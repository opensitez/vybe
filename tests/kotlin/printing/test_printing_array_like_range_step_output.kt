// vybe-test: kotlin/printing/test_printing_array_like_range_step_output
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(3, 6, 9)
            val output = values.joinToString("|")
            __check((output).toString(), "3|6|9")
        }
