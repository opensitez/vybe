// vybe-test: kotlin/throwing_recovery/test_throwing_with_finally_side_effect
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                throw Exception("stop")
            } finally {
                __check(("teardown").toString(), "teardown")
            }
        }
