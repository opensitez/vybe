// vybe-test: kotlin/preconditions/test_check_not_null_message
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                checkNotNull<Int>(null)
            } catch (e: IllegalStateException) {
                __check(("missing").toString(), "missing")
            }
        }
