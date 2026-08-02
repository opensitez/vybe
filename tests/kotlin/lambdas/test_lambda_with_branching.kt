// vybe-test: kotlin/lambdas/test_lambda_with_branching
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val check = { value: Int ->
                if (value > 10) {
                    "big"
                } else {
                    "small"
                }
            }
            __check((check(3)).toString(), "small")
            __check((check(15)).toString(), "big")
        }
