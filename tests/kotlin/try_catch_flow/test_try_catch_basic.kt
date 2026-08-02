// vybe-test: kotlin/try_catch_flow/test_try_catch_basic
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                throw Exception("boom")
            } catch (e: Exception) {
                __check(("caught").toString(), "caught")
            }
        }
