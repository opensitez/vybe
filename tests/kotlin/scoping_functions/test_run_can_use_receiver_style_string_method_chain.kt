// vybe-test: kotlin/scoping_functions/test_run_can_use_receiver_style_string_method_chain
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = "scoping".run {
                uppercase()
            }
            __check((result).toString(), "SCOPING")
        }
