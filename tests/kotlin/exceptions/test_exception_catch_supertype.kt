// vybe-test: kotlin/exceptions/test_exception_catch_supertype
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        throw IllegalArgumentException("bad")
    } catch (e: Exception) {
        __check(("caught").toString(), "caught")
    }
}
