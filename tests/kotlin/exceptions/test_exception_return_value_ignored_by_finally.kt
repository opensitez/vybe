// vybe-test: kotlin/exceptions/test_exception_return_value_ignored_by_finally
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun result(): Int {
    try {
        return 10
    } finally {
        __check(("end").toString(), "end")
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    __check((result()).toString(), "10")
}
