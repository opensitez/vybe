// vybe-test: kotlin/preconditions/test_require_accepts_true_predicate
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            require(true) { "nope" }
            __check(("ok").toString(), "ok")
        }
