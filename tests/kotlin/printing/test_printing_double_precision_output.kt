// vybe-test: kotlin/printing/test_printing_double_precision_output
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3.5).toString(), "3.5")
            __check((0.125).toString(), "0.125")
            __check((-2.0).toString(), "-2.0")
        }
