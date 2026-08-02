// vybe-test: kotlin/literals/test_floating_literals_scientific_notation
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1e3).toString(), "1000")
            __check((2.5e-1).toString(), "0.25")
            __check((3.0E2).toString(), "300")
            __check((1e-3).toString(), "0.001")
        }
