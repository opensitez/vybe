// vybe-test: kotlin/lambdas/test_lambda_in_variable
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sub = { a: Int, b: Int -> a - b }
            __check((sub(100, 30)).toString(), "70")
        }
