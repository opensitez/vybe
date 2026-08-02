// vybe-test: kotlin/preconditions/test_require_true_does_not_throw
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            require(true)
            __check(("ok").toString(), "ok")
        }
