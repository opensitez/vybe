// vybe-test: kotlin/lambdas/test_lambda_try_catch_with_local_state
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val safe = { s: String ->
        try {
            s.toInt()
        } catch (e: Exception) {
            -1
        }
    }
    __check((safe("12")).toString(), "12")
    __check((safe("bad")).toString(), "-1")
}
