// vybe-test: kotlin/lambdas/test_lambda_boolean_check
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val isEven = { x: Int -> x % 2 == 0 }
__check((isEven(8)).toString(), "true")
__check((isEven(5)).toString(), "false") }
