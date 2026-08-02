// vybe-test: kotlin/exceptions/test_exception_finally_around_return_from_nested_context
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun evaluate(): Int {
    return try {
        throw Exception("primary")
    } catch (e: Exception) {
        __check(("inner-catch").toString(), "inner-catch")
        7
    } finally {
        __check(("inner-finally").toString(), "inner-finally")
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    __check((evaluate()).toString(), "7")
}
