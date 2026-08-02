// vybe-test: kotlin/exceptions/test_try_finally_only_with_error
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                throw Exception("panic")
            } finally {
                __check(("ended").toString(), "ended")
            }
        }
