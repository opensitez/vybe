// vybe-test: kotlin/try_catch_flow/test_try_rethrow_in_catch
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                try {
                    throw Exception("x")
                } catch (e: Exception) {
                    throw RuntimeException("wrapped")
                }
            } catch (e: RuntimeException) {
                __check(("wrapped").toString(), "wrapped")
            }
        }
