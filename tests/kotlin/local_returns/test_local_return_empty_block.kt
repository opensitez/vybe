// vybe-test: kotlin/local_returns/test_local_return_empty_block
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = run {
                val v = 0
                if (v == 0) return@run 0
                1
            }
            __check((result).toString(), "0")
        }
