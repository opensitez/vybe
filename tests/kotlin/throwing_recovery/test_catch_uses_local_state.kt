// vybe-test: kotlin/throwing_recovery/test_catch_uses_local_state
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

var seen = 0
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
                seen = 1
            }
            __check((seen).toString(), "1")
        }
