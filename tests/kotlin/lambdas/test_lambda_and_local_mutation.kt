// vybe-test: kotlin/lambdas/test_lambda_and_local_mutation
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0
            val add = { value: Int ->
                total += value
            }
            add(3)
            add(5)
            __check((total).toString(), "8")
        }
