// vybe-test: kotlin/exceptions/test_exception_try_finalize_no_catch
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        __check(("go").toString(), "go")
    } finally {
        __check(("final").toString(), "final")
    }
}
