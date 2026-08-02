// vybe-test: kotlin/lambdas/test_lambda_with_if_expression
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val evaluator = { n: Int ->
                if (n > 10) {
                    "big"
                } else {
                    "small"
                }
            }
            __check((evaluator(2)).toString(), "small")
            __check((evaluator(20)).toString(), "big")
        }
