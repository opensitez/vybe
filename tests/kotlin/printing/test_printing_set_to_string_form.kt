// vybe-test: kotlin/printing/test_printing_set_to_string_form
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = linkedSetOf(1, 2, 3)
            __check((nums).toString(), "[1, 2, 3]")
        }
