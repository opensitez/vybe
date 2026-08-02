// vybe-test: kotlin/printing/test_printing_special_double_values
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((1.0 / 0.0).isInfinite()).toString(), "true")
            __check(((-1.0 / 0.0).isInfinite()).toString(), "true")
            __check(((0.0 / 0.0).isNaN()).toString(), "true")
        }
