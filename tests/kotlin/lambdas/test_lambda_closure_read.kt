// vybe-test: kotlin/lambdas/test_lambda_closure_read
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val factor = 10
            val mult = { x: Int -> x * factor }
            __check((mult(5)).toString(), "50")
        }
