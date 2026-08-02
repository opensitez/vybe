// vybe-test: kotlin/numeric_literals/test_negative_scientific
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((2e-1).toString(), "0.2") }
