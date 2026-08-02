// vybe-test: kotlin/throwing_recovery/test_throwing_in_inline_lambda
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                run {
                    throw Exception("inline")
                }
            } catch (e: Exception) {
                __check((e.message).toString(), "inline")
            }
        }
