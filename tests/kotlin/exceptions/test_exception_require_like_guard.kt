// vybe-test: kotlin/exceptions/test_exception_require_like_guard
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        require(false)
    } catch (e: Exception) {
        __check(("guarded").toString(), "guarded")
    } finally {
        __check(("released").toString(), "released")
    }
}
