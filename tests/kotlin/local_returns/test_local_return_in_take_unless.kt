// vybe-test: kotlin/local_returns/test_local_return_in_take_unless
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = 0
            val out = v.takeUnless { it > 1 } ?: 9
            __check((out).toString(), "9")
        }
