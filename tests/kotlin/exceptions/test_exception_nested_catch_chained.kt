// vybe-test: kotlin/exceptions/test_exception_nested_catch_chained
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        throw Exception("top")
    } catch (e: Exception) {
        try {
            throw IllegalArgumentException("inner")
        } catch (e: IllegalArgumentException) {
            __check(("nested").toString(), "nested")
        }
    }
}
