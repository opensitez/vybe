// vybe-test: kotlin/throwing_recovery/test_throw_and_catch_runtime
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun fail() = throw IllegalStateException("bad")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                fail()
            } catch (e: IllegalStateException) {
                __check((e.message).toString(), "bad")
            }
        }
