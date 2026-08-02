// vybe-test: kotlin/scope/test_scope_try_finally_alters_and_observes_outer_mutation
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var state = "open"
            try {
                state = "processing"
            } finally {
                state = state + "-done"
            }
            __check((state).toString(), "processing-done")
        }
