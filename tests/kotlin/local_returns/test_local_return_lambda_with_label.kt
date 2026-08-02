// vybe-test: kotlin/local_returns/test_local_return_lambda_with_label
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = listOf(1, 2, 3).filterIndexed { index, value ->
                if (index == 1) return@filterIndexed false
                value % 2 == 1
            }
            __check((out.joinToString(",")).toString(), "1,3")
        }
