// vybe-test: kotlin/printing/test_printing_zero_and_sign_boundary
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0).toString(), "0")
            __check((+5).toString(), "5")
            __check((-0).toString(), "0")
        }
