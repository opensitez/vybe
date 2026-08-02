// vybe-test: kotlin/printing/test_printing_dollar_sign_escape
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("cost: ${'$'}5").toString(), "cost: \$5")
        }
