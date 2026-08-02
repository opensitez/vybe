// vybe-test: kotlin/lambdas/test_lambda_reference_reuse_and_reassignment
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    var op: (Int) -> Int = { it + 1 }
    __check((op(4)).toString(), "5")
    op = { it * 2 }
    __check((op(4)).toString(), "8")
}
