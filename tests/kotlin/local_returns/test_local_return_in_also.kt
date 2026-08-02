// vybe-test: kotlin/local_returns/test_local_return_in_also
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = 4
            val out = n.also {
                if (it < 0) return
            }
            __check((out).toString(), "4")
        }
