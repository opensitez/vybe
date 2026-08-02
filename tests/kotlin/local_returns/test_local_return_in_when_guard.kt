// vybe-test: kotlin/local_returns/test_local_return_in_when_guard
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 5
            val label = when {
                value > 10 -> "high"
                value < 0 -> "low"
                else -> run { value.toString() }
            }
            __check((label).toString(), "5")
        }
