// vybe-test: kotlin/try_finally/test_finally_always_runs_even_if_caught
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run() {
        try {
            try {
                throw RuntimeException()
            } finally {
                __check(("inner").toString(), "inner")
            }
        } catch (e: Exception) {
            __check(("outer").toString(), "outer")
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { run() }
