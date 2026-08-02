// vybe-test: kotlin/exceptions/test_exception_throw_in_finally_overrides_body_exception
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
            throw Exception("body")
        } finally {
            __check(("body-finally").toString(), "body-finally")
            throw Exception("finally")
        }
    } catch (e: Exception) {
        __check((e.message).toString(), "finally")
    }
}
