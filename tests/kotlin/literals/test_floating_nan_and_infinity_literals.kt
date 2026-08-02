// vybe-test: kotlin/literals/test_floating_nan_and_infinity_literals
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val zero = 0.0
            val nan = 0.0 / zero
            val inf = 1.0 / zero
            __check((nan.isNaN()).toString(), "true")
            __check((inf > 0).toString(), "true")
        }
