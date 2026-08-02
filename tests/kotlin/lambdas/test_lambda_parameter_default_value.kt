// vybe-test: kotlin/lambdas/test_lambda_parameter_default_value
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val add = { base: Int, bonus: Int -> base + bonus }
__check((add(6, 4)).toString(), "10") }
