// vybe-test: kotlin/apply_scope_functions/test_run_and_with_return_different_receiver
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val withValue = with("a") { this + "bc" }
            val runValue = "a".run { uppercase() + "bc" }
            __check((withValue).toString(), "abc")
            __check((runValue).toString(), "Abc")
        }
