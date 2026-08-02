// vybe-test: kotlin/exceptions/test_exception_resource_like_flow
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun run() {
    try {
        __check(("open").toString(), "open")
    } finally {
        __check(("closed").toString(), "closed")
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    run()
}
