// vybe-test: kotlin/local_returns/test_local_return_in_with
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = with(0) {
                if (this < 0) return@with 0
                this + 1
            }
            __check((value).toString(), "1")
        }
