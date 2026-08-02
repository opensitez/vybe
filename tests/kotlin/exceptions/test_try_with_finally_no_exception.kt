// vybe-test: kotlin/exceptions/test_try_with_finally_no_exception
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                __check(("ok").toString(), "ok")
            } finally {
                __check(("cleanup").toString(), "cleanup")
            }
        }
