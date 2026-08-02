// vybe-test: kotlin/try_finally/test_try_with_exception_caught
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        try {
            throw IllegalStateException("x")
        } catch (e: Exception) {
            __check((e::class.simpleName).toString(), "IllegalStateException")
        } finally {
            __check(("done").toString(), "done")
        }
    }
