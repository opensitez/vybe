// vybe-test: kotlin/throwing_recovery/test_throwable_chain_and_recover
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

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
                try {
                    __check(("inner").toString(), "inner")
                    throw RuntimeException("y")
                } catch (inner: RuntimeException) {
                    __check((inner.message).toString(), "y")
                }
            }
        }
