// vybe-test: kotlin/lambdas/test_lambda_with_multiple_parameters_destructuring_syntax
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val merge = { (left, right): Pair<Int, Int> ->
        left + right
    }
    __check((merge(Pair(7, 8))).toString(), "15")
}
