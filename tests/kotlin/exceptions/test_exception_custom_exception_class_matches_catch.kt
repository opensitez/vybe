// vybe-test: kotlin/exceptions/test_exception_custom_exception_class_matches_catch
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

class NetworkError(message: String) : Exception(message)

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        throw NetworkError("down")
    } catch (e: NetworkError) {
        __check(("custom").toString(), "custom")
        __check((e.message).toString(), "down")
    }
}
