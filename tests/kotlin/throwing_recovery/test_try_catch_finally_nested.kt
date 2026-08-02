// vybe-test: kotlin/throwing_recovery/test_try_catch_finally_nested
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
                    throw Exception("inner")
                } catch (e: Exception) {
                    __check(("inner").toString(), "inner")
                } finally {
                    __check(("inner finally").toString(), "inner finally")
                }
            } finally {
                __check(("outer finally").toString(), "outer finally")
            }
        }
