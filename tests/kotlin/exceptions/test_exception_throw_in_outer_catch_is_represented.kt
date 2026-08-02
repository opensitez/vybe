// vybe-test: kotlin/exceptions/test_exception_throw_in_outer_catch_is_represented
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
            throw Exception("outer")
        }
    } catch (e: Exception) {
        __check(("caught").toString(), "caught")
        __check((e.message).toString(), "outer")
    }
}
