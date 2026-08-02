// vybe-test: kotlin/throwing_recovery/test_throwing_after_catch_cleanup
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

var cleaned = 0
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
                } finally {
                    cleaned += 1
                }
            } catch (e: Exception) {
                __check((cleaned).toString(), "1")
            }
        }
