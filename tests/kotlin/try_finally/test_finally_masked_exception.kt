// vybe-test: kotlin/try_finally/test_finally_masked_exception
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun probe(): Int {
        try {
            try {
                throw IllegalStateException("x")
            } finally {
                throw RuntimeException("y")
            }
        } catch (e: RuntimeException) {
            __check((e.message).toString(), "y")
            return 0
        } finally {
            __check(("after").toString(), "after")
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((probe()).toString(), "0") }
