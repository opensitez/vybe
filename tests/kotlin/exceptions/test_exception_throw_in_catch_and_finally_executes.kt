// vybe-test: kotlin/exceptions/test_exception_throw_in_catch_and_finally_executes
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
            throw Exception("inner")
        } catch (e: Exception) {
            __check(("caught-inner").toString(), "caught-inner")
            throw Exception("from-catch")
        } finally {
            __check(("inner-finally").toString(), "inner-finally")
        }
    } catch (e: Exception) {
        __check(("caught-outer").toString(), "caught-outer")
        __check((e.message).toString(), "from-catch")
    }
}
