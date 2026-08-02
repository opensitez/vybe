// vybe-test: kotlin/throwing_recovery/test_rethrow_new_exception
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                try {
                    throw Exception("a")
                } catch (e: Exception) {
                    throw RuntimeException("b")
                }
            } catch (e: RuntimeException) {
                __check((e.message).toString(), "b")
            }
        }
