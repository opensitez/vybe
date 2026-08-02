// vybe-test: kotlin/lambdas/test_lambda_capture_and_modify
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var base = 1
val inc = { x: Int -> base + x }
__check((inc(4)).toString(), "5") }
