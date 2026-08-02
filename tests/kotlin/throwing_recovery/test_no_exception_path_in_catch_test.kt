// vybe-test: kotlin/throwing_recovery/test_no_exception_path_in_catch_test
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun mayFail(shouldFail: Boolean): Int {
            return try {
                if (shouldFail) throw Exception("x")
                1
            } catch (e: Exception) {
                0
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((mayFail(false)).toString(), "1")
            __check((mayFail(true)).toString(), "0")
        }
