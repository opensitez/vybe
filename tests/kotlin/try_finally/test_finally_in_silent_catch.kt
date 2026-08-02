// vybe-test: kotlin/try_finally/test_finally_in_silent_catch
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run() {
        try {
            throw IllegalArgumentException()
        } catch (e: RuntimeException) {
            // ignore
        } finally {
            __check(("done").toString(), "done")
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { run() }
