// vybe-test: kotlin/exceptions/test_exception_require_and_finally_cleanup
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        try {
            throw Exception("boom")
        } finally {
            __check(("cleanup").toString(), "cleanup")
        }
    } catch (e: Exception) {
        __check(("caught").toString(), "caught")
    }
}
