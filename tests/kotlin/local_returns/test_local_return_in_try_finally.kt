// vybe-test: kotlin/local_returns/test_local_return_in_try_finally
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = run {
                try {
                    return@run "ok"
                } finally {
                    __check(("fin").toString(), "fin")
                }
            }
            __check((out).toString(), "ok")
        }
