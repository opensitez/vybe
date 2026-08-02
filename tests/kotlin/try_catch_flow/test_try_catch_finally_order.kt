// vybe-test: kotlin/try_catch_flow/test_try_catch_finally_order
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                throw Exception("x")
            } catch (e: Exception) {
                __check(("catch").toString(), "catch")
            } finally {
                __check(("finally").toString(), "finally")
            }
            __check(("after").toString(), "after")
        }
