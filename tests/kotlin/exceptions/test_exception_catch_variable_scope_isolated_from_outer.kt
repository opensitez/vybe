// vybe-test: kotlin/exceptions/test_exception_catch_variable_scope_isolated_from_outer
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val message = "root"
    try {
        throw Exception("inner")
    } catch (message: Exception) {
        __check((message.message).toString(), "inner")
    }
    __check(("root").toString(), "root")
}
