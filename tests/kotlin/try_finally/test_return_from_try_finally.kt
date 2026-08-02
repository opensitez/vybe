// vybe-test: kotlin/try_finally/test_return_from_try_finally
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun probe(): String {
        try {
            return "try"
        } finally {
            return "finally"
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((probe()).toString(), "finally") }
