// vybe-test: kotlin/exceptions/test_exception_try_in_function
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun risky(value: Int) {
    if (value < 0) {
        throw Exception("no")
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        risky(-1)
    } catch (e: Exception) {
        __check(("blocked").toString(), "blocked")
    }
}
