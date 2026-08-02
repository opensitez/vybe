// vybe-test: kotlin/scoping_functions/test_run_can_catch_and_continue_from_exception_in_scope
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = try {
                run {
                    throw RuntimeException("boom")
                }
            } catch (error: RuntimeException) {
                "caught"
            }
            __check((out).toString(), "caught")
        }
