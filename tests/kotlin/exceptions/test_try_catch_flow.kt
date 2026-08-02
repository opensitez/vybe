// vybe-test: kotlin/exceptions/test_try_catch_flow
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
                throw Exception("failure")
            } catch (e: Exception) {
                __check(("catch block").toString(), "catch block")
            }
        }
