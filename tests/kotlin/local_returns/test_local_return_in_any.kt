// vybe-test: kotlin/local_returns/test_local_return_in_any
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val anyEven = listOf(1, 3, 4).any {
                if (it == 2) return@any false
                it % 2 == 0
            }
            __check((anyEven).toString(), "true")
        }
