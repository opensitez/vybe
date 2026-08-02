// vybe-test: kotlin/local_returns/test_local_return_from_try
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
                    throw Exception("x")
                } catch (e: Exception) {
                    return@run "err"
                }
                "ok"
            }
            __check((out).toString(), "err")
        }
