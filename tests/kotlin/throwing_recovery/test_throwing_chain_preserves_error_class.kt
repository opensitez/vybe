// vybe-test: kotlin/throwing_recovery/test_throwing_chain_preserves_error_class
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
                    throw IllegalArgumentException("bad")
                } catch (e: Exception) {
                    throw RuntimeException(e)
                }
            } catch (e: RuntimeException) {
                __check((e.cause != null).toString(), "true")
            }
        }
