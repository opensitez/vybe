// vybe-test: kotlin/try_catch_flow/test_catch_only_finally
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
            } finally {
                __check(("ok").toString(), "ok")
            }
        }
