// vybe-test: kotlin/printing/test_printing_boolean_values
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true).toString(), "true")
            __check((false).toString(), "false")
            __check((1 == 1).toString(), "true")
            __check((1 == 2).toString(), "false")
        }
