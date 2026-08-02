// vybe-test: kotlin/exceptions/test_exception_nested_finally_order
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
            __check(("inner").toString(), "inner")
        } finally {
            __check(("inner finally").toString(), "inner finally")
        }
    } finally {
        __check(("outer finally").toString(), "outer finally")
    }
}
