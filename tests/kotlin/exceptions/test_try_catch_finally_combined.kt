// vybe-test: kotlin/exceptions/test_try_catch_finally_combined
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                __check(("working").toString(), "working")
                throw Exception("bad arg")
            } catch (e: Exception) {
                __check(("handled arg error").toString(), "handled arg error")
            } finally {
                __check(("cleanup done").toString(), "cleanup done")
            }
        }
