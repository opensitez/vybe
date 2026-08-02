// vybe-test: kotlin/lambdas/test_lambda_returning_boolean
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val isGreater = { a: Int, b: Int -> a > b }
            __check((isGreater(10, 5)).toString(), "true")
            __check((isGreater(2, 8)).toString(), "false")
        }
