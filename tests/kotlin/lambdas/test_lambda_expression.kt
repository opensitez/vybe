// vybe-test: kotlin/lambdas/test_lambda_expression
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mult = { a: Int, b: Int -> a * b }
            __check((mult(4, 5)).toString(), "20")
        }
