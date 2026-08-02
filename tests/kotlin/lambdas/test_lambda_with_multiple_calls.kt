// vybe-test: kotlin/lambdas/test_lambda_with_multiple_calls
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun runTwice(v: Int, op: (Int) -> Int): Int { return op(op(v)) }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((runTwice(2, { it * 3 })).toString(), "18") }
