// vybe-test: kotlin/local_returns/test_local_return_in_fold
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val total = listOf(1, 2, 3).fold(0) { acc, value ->
                if (value == 2) return@fold acc
                acc + value
            }
            __check((total).toString(), "4")
        }
