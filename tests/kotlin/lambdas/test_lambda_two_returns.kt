// vybe-test: kotlin/lambdas/test_lambda_two_returns
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val choose = { flag: Boolean -> if (flag) "yes" else "no" }
__check((choose(true)).toString(), "yes")
__check((choose(false)).toString(), "no") }
