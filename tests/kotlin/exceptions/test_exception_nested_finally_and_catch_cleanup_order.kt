// vybe-test: kotlin/exceptions/test_exception_nested_finally_and_catch_cleanup_order
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
            __check(("inner-try").toString(), "inner-try")
            throw Exception("inner")
        } catch (e: Exception) {
            __check(("inner-catch").toString(), "inner-catch")
            throw e
        } finally {
            __check(("inner-finally").toString(), "inner-finally")
        }
    } catch (e: Exception) {
        __check(("outer-catch").toString(), "outer-catch")
    } finally {
        __check(("outer-finally").toString(), "outer-finally")
    }
}
