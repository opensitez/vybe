// vybe-test: kotlin/lambdas/test_lambda_with_no_arguments
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val supplier = { "value" }
            __check((supplier()).toString(), "value")
        }
