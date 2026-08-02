// vybe-test: kotlin/lambdas/test_lambda_as_expression_argument_without_params
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun runTask(task: () -> String): String {
            return task()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((runTask({ "done" })).toString(), "done")
        }
