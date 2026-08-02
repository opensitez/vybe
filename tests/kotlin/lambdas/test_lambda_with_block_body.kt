// vybe-test: kotlin/lambdas/test_lambda_with_block_body
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val compute = { value: Int ->
                val x = value * 2
                val y = x + 1
                y
            }
            __check((compute(5)).toString(), "11")
        }
