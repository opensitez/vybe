// vybe-test: kotlin/local_returns/test_local_return_in_all
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val allPositive = listOf(1, 2, -1).all {
                if (it < 0) return@all false
                it > 0
            }
            __check((allPositive).toString(), "false")
        }
