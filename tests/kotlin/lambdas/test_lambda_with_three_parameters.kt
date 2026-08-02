// vybe-test: kotlin/lambdas/test_lambda_with_three_parameters
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val combine = { a: Int, b: Int, c: Int -> a + b + c }
            __check((combine(1, 2, 3)).toString(), "6")
        }
