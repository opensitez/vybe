// vybe-test: kotlin/try_finally/test_finally_after_throw
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        try {
            try {
                throw RuntimeException("boom")
            } finally {
                __check(("clean").toString(), "clean")
            }
        } catch (e: Exception) {
            __check(("caught").toString(), "caught")
        }
    }
