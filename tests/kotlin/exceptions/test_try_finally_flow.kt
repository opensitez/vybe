// vybe-test: kotlin/exceptions/test_try_finally_flow
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                __check(("try start").toString(), "try start")
            } finally {
                __check(("finally block").toString(), "finally block")
            }
        }
