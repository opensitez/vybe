// vybe-test: kotlin/lambdas/test_lambda_stored_in_array
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val ops = arrayOf({ x: Int -> x + 1 }, { x: Int -> x * 2 })
__check((ops[0](3)).toString(), "4")
__check((ops[1](3)).toString(), "6") }
